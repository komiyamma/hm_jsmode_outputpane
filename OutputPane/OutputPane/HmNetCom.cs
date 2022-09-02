/*
 * HmNetCOM ver 2.071
 * Copyright (C) 2021-2022 Akitsugu Komiyama
 * under the MIT License
 **/

using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace HmNetCOM
{
    internal partial class Hm
    {
        public static partial class OutputPane
        {
            private static UnManagedDll hmOutputPaneHandle = null;

            // OutputPaneから出ている関数群
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate int TOutputPane_Output(IntPtr hHidemaruWindow, byte[] encode_data);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate int TOutputPane_OutputW(IntPtr hHidemaruWindow, [MarshalAs(UnmanagedType.LPWStr)] String pwszmsg);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate int TOutputPane_Push(IntPtr hHidemaruWindow);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate int TOutputPane_Pop(IntPtr hHidemaruWindow);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate IntPtr TOutputPane_GetWindowHandle(IntPtr hHidemaruWindow);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate int TOutputPane_SetBaseDir(IntPtr hHidemaruWindow, byte[] encode_data);

            private static TOutputPane_Output pOutputPane_Output;
            private static TOutputPane_OutputW pOutputPane_OutputW;
            private static TOutputPane_Push pOutputPane_Push;
            private static TOutputPane_Pop pOutputPane_Pop;
            private static TOutputPane_GetWindowHandle pOutputPane_GetWindowHandle;
            private static TOutputPane_SetBaseDir pOutputPane_SetBaseDir;

            static OutputPane()
            {
                try
                {
                    string exedir = System.IO.Path.GetDirectoryName(GetHidemaruExeFullPath());
                    hmOutputPaneHandle = new UnManagedDll(Path.Combine(exedir, "HmOutputPane.dll"));
                    pOutputPane_Output = hmOutputPaneHandle.GetProcDelegate<TOutputPane_Output>("Output");
                    pOutputPane_Push = hmOutputPaneHandle.GetProcDelegate<TOutputPane_Push>("Push");
                    pOutputPane_Pop = hmOutputPaneHandle.GetProcDelegate<TOutputPane_Pop>("Pop");
                    pOutputPane_GetWindowHandle = hmOutputPaneHandle.GetProcDelegate<TOutputPane_GetWindowHandle>("GetWindowHandle");

                    if (Version >= 877)
                    {
                        pOutputPane_SetBaseDir = hmOutputPaneHandle.GetProcDelegate<TOutputPane_SetBaseDir>("SetBaseDir");
                    }
                    if (Version >= 898)
                    {
                        pOutputPane_OutputW = hmOutputPaneHandle.GetProcDelegate<TOutputPane_OutputW>("OutputW");
                    }
                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message);
                }
            }

            /// <summary>
            /// アウトプット枠への文字列の出力。
            /// 改行するには「\r\n」といったように「\r」も必要。
            /// </summary>
            /// <returns>失敗なら0、成功なら0以外</returns>
            public static int Output(string message)
            {
                try
                {
                    if (pOutputPane_OutputW != null)
                    {
                        int result = pOutputPane_OutputW(Hm.WindowHandle, message);
                        return result;
                    }
                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message);
                }

                return 0;
            }

            /// <summary>
            /// アウトプット枠にある文字列の一時退避
            /// </summary>
            /// <returns>失敗なら0、成功なら0以外</returns>
            public static int Push()
            {
                return pOutputPane_Push(Hm.WindowHandle); ;
            }

            /// <summary>
            /// Pushによって一時退避した文字列の復元
            /// </summary>
            /// <returns>失敗なら0、成功なら0以外</returns>
            public static int Pop()
            {
                return pOutputPane_Pop(Hm.WindowHandle); ;
            }

            /// <summary>
            /// アウトプット枠にある文字列のクリア
            /// </summary>
            /// <returns>現在のところ、成否を指し示す値は返ってこない</returns>
            public static int Clear()
            {
                //1009=クリア
                IntPtr r = OutputPane.SendMessage(1009);
                if ((long)r < (long)int.MinValue)
                {
                    r = (IntPtr)int.MinValue;
                }
                if ((long)int.MaxValue < (long)r)
                {
                    r = (IntPtr)int.MaxValue;
                }

                return (int)r;
            }

            /// <summary>
            /// アウトプット枠のWindowHandle
            /// </summary>
            /// <returns>アウトプット枠のWindowHandle</returns>
            public static IntPtr WindowHandle
            {
                get
                {
                    return pOutputPane_GetWindowHandle(Hm.WindowHandle);
                }
            }

            /// <summary>
            /// アウトプット枠へのSendMessage
            /// </summary>
            /// <returns>SendMessageの返り値そのまま</returns>
            public static IntPtr SendMessage(int commandID)
            {
                IntPtr result = Hm.SendMessage(OutputPane.WindowHandle, 0x111, (IntPtr)commandID, IntPtr.Zero);
                return result;
            }

        }
    }
}

namespace HmNetCOM
{
    internal partial class Hm
    {
        public static partial class ExplorerPane
        {
            private static UnManagedDll hmExplorerPaneHandle = null;

            // ExplorerPaneから出ている関数群
            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate int TExplorerPane_SetMode(IntPtr hHidemaruWindow, IntPtr mode);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate int TExplorerPane_GetMode(IntPtr hHidemaruWindow);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate int TExplorerPane_LoadProject(IntPtr hHidemaruWindow, byte[] encode_project_file_path);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate int TExplorerPane_SaveProject(IntPtr hHidemaruWindow, byte[] encode_project_file_path);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate IntPtr TExplorerPane_GetProject(IntPtr hHidemaruWindow);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate IntPtr TExplorerPane_GetWindowHandle(IntPtr hHidemaruWindow);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate int TExplorerPane_GetUpdated(IntPtr hHidemaruWindow);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            private delegate IntPtr TExplorerPane_GetCurrentDir(IntPtr hHidemaruWindow);

            private static TExplorerPane_SetMode pExplorerPane_SetMode;
            private static TExplorerPane_GetMode pExplorerPane_GetMode;
            private static TExplorerPane_LoadProject pExplorerPane_LoadProject;
            private static TExplorerPane_SaveProject pExplorerPane_SaveProject;
            private static TExplorerPane_GetProject pExplorerPane_GetProject;
            private static TExplorerPane_GetWindowHandle pExplorerPane_GetWindowHandle;
            private static TExplorerPane_GetUpdated pExplorerPane_GetUpdated;
            private static TExplorerPane_GetCurrentDir pExplorerPane_GetCurrentDir;

            static ExplorerPane()
            {
                try
                {
                    string exedir = System.IO.Path.GetDirectoryName(GetHidemaruExeFullPath());
                    hmExplorerPaneHandle = new UnManagedDll(exedir + @"\HmExplorerPane.dll");
                    pExplorerPane_SetMode = hmExplorerPaneHandle.GetProcDelegate<TExplorerPane_SetMode>("SetMode");
                    pExplorerPane_GetMode = hmExplorerPaneHandle.GetProcDelegate<TExplorerPane_GetMode>("GetMode");
                    pExplorerPane_LoadProject = hmExplorerPaneHandle.GetProcDelegate<TExplorerPane_LoadProject>("LoadProject");
                    pExplorerPane_SaveProject = hmExplorerPaneHandle.GetProcDelegate<TExplorerPane_SaveProject>("SaveProject");
                    pExplorerPane_GetProject = hmExplorerPaneHandle.GetProcDelegate<TExplorerPane_GetProject>("GetProject");
                    pExplorerPane_GetUpdated = hmExplorerPaneHandle.GetProcDelegate<TExplorerPane_GetUpdated>("GetUpdated");
                    pExplorerPane_GetWindowHandle = hmExplorerPaneHandle.GetProcDelegate<TExplorerPane_GetWindowHandle>("GetWindowHandle");

                    if (Version >= 885)
                    {
                        pExplorerPane_GetCurrentDir = hmExplorerPaneHandle.GetProcDelegate<TExplorerPane_GetCurrentDir>("GetCurrentDir");
                    }
                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message);
                }
            }

            /// <summary>
            /// ファイルマネージャ枠のモードの設定
            /// </summary>
            /// <returns>失敗なら0、成功なら0以外</returns>
            public static int SetMode(int mode)
            {
                try
                {
                    int result = pExplorerPane_SetMode(Hm.WindowHandle, (IntPtr)mode);
                    return result;
                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message);
                }

                return 0;
            }

            /// <summary>
            /// ファイルマネージャ枠のモードの取得
            /// </summary>
            /// <returns>モードの値</returns>
            public static int GetMode()
            {
                try
                {
                    int result = pExplorerPane_GetMode(Hm.WindowHandle);
                    return result;
                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message);
                }

                return 0;
            }

            /// <summary>
            /// ファイルマネージャ枠にプロジェクトを読み込んでいるならば、そのファイルパスを取得する
            /// </summary>
            /// <returns>ファイルのフルパス。読み込んでいなければnull</returns>
            public static string GetProject()
            {
                try
                {
                    IntPtr startpointer = pExplorerPane_GetProject(Hm.WindowHandle);
                    List<byte> blist = GetPointerToByteArray(startpointer);

                    string project_name = HmOriginalDecodeFunc.DecodeOriginalEncodeVector(blist);

                    if (String.IsNullOrEmpty(project_name))
                    {
                        return null;
                    }
                    return project_name;
                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message);
                }

                return null;
            }

            private static List<byte> GetPointerToByteArray(IntPtr startpointer)
            {
                List<byte> blist = new List<byte>();

                int index = 0;
                while (true)
                {
                    var b = Marshal.ReadByte(startpointer, index);

                    blist.Add(b);

                    // 文字列の終端はやはり0
                    if (b == 0)
                    {
                        break;
                    }

                    index++;
                }

                return blist;
            }

            /// <summary>
            /// ファイルマネージャ枠のカレントディレクトリを返す
            /// </summary>
            /// <returns>カレントディレクトリのフルパス。読み損ねた場合はnull</returns>
            public static string GetCurrentDir()
            {
                if (Version < 885)
                {
                    throw new MissingMethodException("HmOutputPane_GetCurrentDir_Exception");
                }
                try
                {
                    if (pExplorerPane_GetCurrentDir != null)
                    {
                        IntPtr startpointer = pExplorerPane_GetCurrentDir(Hm.WindowHandle);
                        List<byte> blist = GetPointerToByteArray(startpointer);

                        string currentdir_name = HmOriginalDecodeFunc.DecodeOriginalEncodeVector(blist);

                        if (String.IsNullOrEmpty(currentdir_name))
                        {
                            return null;
                        }
                        return currentdir_name;
                    }
                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message);
                }

                return null;
            }

            /// <summary>
            /// ファイルマネージャ枠が「プロジェクト」表示のとき、更新された状態であるかどうかを返します
            /// </summary>
            /// <returns>更新状態なら1、それ以外は0</returns>
            public static int GetUpdated()
            {
                try
                {
                    int result = pExplorerPane_GetUpdated(Hm.WindowHandle);
                    return result;
                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message);
                }

                return 0;
            }

            /// <summary>
            /// ファイルマネージャ枠のWindowHandle
            /// </summary>
            /// <returns>ファイルマネージャ枠のWindowHandle</returns>
            public static IntPtr WindowHandle
            {
                get
                {
                    return pExplorerPane_GetWindowHandle(Hm.WindowHandle);
                }
            }

            /// <summary>
            /// ファイルマネージャ枠へのSendMessage
            /// </summary>
            /// <returns>SendMessageの返り値そのまま</returns>
            public static IntPtr SendMessage(int commandID)
            {
                //
                // loaddll "HmExplorerPane.dll";
                // #h=dllfunc("GetWindowHandle",hidemaruhandle(0));
                // #ret=sendmessage(#h,0x111/*WM_COMMAND*/,251,0); //251=１つ上のフォルダ
                //
                return Hm.SendMessage(ExplorerPane.WindowHandle, 0x111, (IntPtr)commandID, IntPtr.Zero);
            }

        }
    }
}


namespace HmNetCOM
{
    internal partial class Hm
    {
        internal static partial class HmOriginalDecodeFunc
        {
            static bool IsSTARTUNI_inline(uint byte4)
            {
                return (byte4 & 0xF4808000) == 0x04808000;
            }

            static long MakeWord(long low, long high)
            {
                return ((long)high << 8) | low;
            }

            static char GetUnicodeInText(byte[] pchSrc)
            {
                long value = MakeWord(
                    (pchSrc[1] & 0x7F | ((pchSrc[3] & 0x01) << 7)),
                    (pchSrc[2] & 0x7F | ((pchSrc[3] & 0x02) << 6))
                );

                byte[] byteArray = BitConverter.GetBytes(value);

                byte[] charByte = { byteArray[0], byteArray[1] };

                char wch = BitConverter.ToChar(charByte, 0);

                return wch;
            }

            public static string DecodeOriginalEncodeVector(List<byte> OriginalEncodeData)
            {
                try
                {
                    string result = "";

                    byte[] byteArray = OriginalEncodeData.ToArray();

                    // 一時バッファー用
                    List<byte> tmp_buffer = new List<byte>();
                    int len = OriginalEncodeData.Count;

                    int lastcheckindex = len - 4; // IsSTARTUNI_inline には 4バイト必要
                    if (lastcheckindex < 0)
                    {
                        lastcheckindex = 0;
                    }
                    for (int i = 0; i < len; i++)
                    {
                        // 一般の文字としてはほぼ利用されないスターマーク。
                        if (i <= lastcheckindex && byteArray[i] == '\x1A')
                        {
                            uint StarUni = BitConverter.ToUInt32(byteArray, i);

                            if (IsSTARTUNI_inline(StarUni))
                            {
                                // 今までの分はスターユニコードではないので、通常のSJISとみなし、utf16に変換して足し込み
                                if (tmp_buffer.Count > 0)
                                {
                                    result += System.Text.Encoding.GetEncoding(932).GetString(tmp_buffer.ToArray());
                                    tmp_buffer.Clear();
                                }

                                byte[] starByteArray = BitConverter.GetBytes(StarUni);
                                char wch = GetUnicodeInText(starByteArray);
                                i = i + 3; // 1バイトではなく4バイト消化したので、計算する
                                result += wch;
                                continue;
                            }
                        }
                        tmp_buffer.Add(byteArray[i]);
                    }

                    if (tmp_buffer.Count > 0)
                    {
                        result += System.Text.Encoding.GetEncoding(932).GetString(tmp_buffer.ToArray());
                        tmp_buffer.Clear();
                    }

                    return result;
                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e);
                }

                return "";
            }
        }

    }
}





/*
 * Copyright (C) 2021-2022 Akitsugu Komiyama
 * under the MIT License
 **/


namespace HmNetCOM
{
    internal partial class Hm
    {
        public interface IComDetachMethod
        {
            void OnReleaseObject(int reason = 0);
        }

        public interface IComSupportX64
        {
#if (NET || NETCOREAPP3_1)

            bool X64MACRO() { return true; }
#else
            bool X64MACRO();
#endif
        }

        static Hm()
        {
            SetVersion();
            BindHidemaruExternFunctions();
        }

        private static void SetVersion()
        {
            if (Version == 0)
            {
                string hidemaru_fullpath = GetHidemaruExeFullPath();
                System.Diagnostics.FileVersionInfo vi = System.Diagnostics.FileVersionInfo.GetVersionInfo(hidemaru_fullpath);
                Version = 100 * vi.FileMajorPart + 10 * vi.FileMinorPart + 1 * vi.FileBuildPart + 0.01 * vi.FilePrivatePart;
            }
        }
        /// <summary>
        /// 秀丸バージョンの取得
        /// </summary>
        /// <returns>秀丸バージョン</returns>
        public static double Version { get; private set; } = 0;

        private const int filePathMaxLength = 512;

        private static string GetHidemaruExeFullPath()
        {
            var sb = new StringBuilder(filePathMaxLength);
            GetModuleFileName(IntPtr.Zero, sb, filePathMaxLength);
            string hidemaru_fullpath = sb.ToString();
            return hidemaru_fullpath;
        }

        /// <summary>
        /// 呼ばれたプロセスの現在の秀丸エディタのウィンドウハンドルを返します。
        /// </summary>
        /// <returns>現在の秀丸エディタのウィンドウハンドル</returns>
        public static IntPtr WindowHandle
        {
            get
            {
                return pGetCurrentWindowHandle();
            }
        }

        private static T HmClamp<T>(T val, T min, T max) where T : IComparable<T>
        {
            if (val.CompareTo(min) < 0) return min;
            else if (val.CompareTo(max) > 0) return max;
            else return val;
        }

        private static bool IsDoubleNumeric(object value)
        {
            return value is double || value is float;
        }
    }
}

namespace HmNetCOM
{

    internal partial class Hm
    {
        // 秀丸本体から出ている関数群
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TGetCurrentWindowHandle();

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TGetTotalTextUnicode();

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TGetLineTextUnicode(int nLineNo);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TGetSelectedTextUnicode();

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TGetCursorPosUnicode(out int pnLineNo, out int pnColumn);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TGetCursorPosUnicodeFromMousePos(IntPtr lpPoint, out int pnLineNo, out int pnColumn);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TEvalMacro([MarshalAs(UnmanagedType.LPWStr)] String pwsz);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TCheckQueueStatus();

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TAnalyzeEncoding([MarshalAs(UnmanagedType.LPWStr)] String pwszFileName, IntPtr lParam1, IntPtr lParam2);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TLoadFileUnicode([MarshalAs(UnmanagedType.LPWStr)] String pwszFileName, int nEncode, ref int pcwchOut, IntPtr lParam1, IntPtr lParam2);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TGetStaticVariable([MarshalAs(UnmanagedType.LPWStr)] String pwszSymbolName, int sharedMemoryFlag);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TSetStaticVariable([MarshalAs(UnmanagedType.LPWStr)] String pwszSymbolName, [MarshalAs(UnmanagedType.LPWStr)] String pwszValue, int sharedMemoryFlag);

        // 秀丸本体から出ている関数群
        private static TGetCurrentWindowHandle pGetCurrentWindowHandle;
        private static TGetTotalTextUnicode pGetTotalTextUnicode;
        private static TGetLineTextUnicode pGetLineTextUnicode;
        private static TGetSelectedTextUnicode pGetSelectedTextUnicode;
        private static TGetCursorPosUnicode pGetCursorPosUnicode;
        private static TGetCursorPosUnicodeFromMousePos pGetCursorPosUnicodeFromMousePos;
        private static TEvalMacro pEvalMacro;
        private static TCheckQueueStatus pCheckQueueStatus;
        private static TAnalyzeEncoding pAnalyzeEncoding;
        private static TLoadFileUnicode pLoadFileUnicode;
        private static TGetStaticVariable pGetStaticVariable;
        private static TSetStaticVariable pSetStaticVariable;

        // 秀丸本体のexeを指すモジュールハンドル
        private static UnManagedDll hmExeHandle;

        private static void BindHidemaruExternFunctions()
        {
            // 初めての代入のみ
            if (hmExeHandle == null)
            {
                try
                {
                    hmExeHandle = new UnManagedDll(GetHidemaruExeFullPath());

                    pGetTotalTextUnicode = hmExeHandle.GetProcDelegate<TGetTotalTextUnicode>("Hidemaru_GetTotalTextUnicode");
                    pGetLineTextUnicode = hmExeHandle.GetProcDelegate<TGetLineTextUnicode>("Hidemaru_GetLineTextUnicode");
                    pGetSelectedTextUnicode = hmExeHandle.GetProcDelegate<TGetSelectedTextUnicode>("Hidemaru_GetSelectedTextUnicode");
                    pGetCursorPosUnicode = hmExeHandle.GetProcDelegate<TGetCursorPosUnicode>("Hidemaru_GetCursorPosUnicode");
                    pEvalMacro = hmExeHandle.GetProcDelegate<TEvalMacro>("Hidemaru_EvalMacro");
                    pCheckQueueStatus = hmExeHandle.GetProcDelegate<TCheckQueueStatus>("Hidemaru_CheckQueueStatus");

                    pGetCursorPosUnicodeFromMousePos = hmExeHandle.GetProcDelegate<TGetCursorPosUnicodeFromMousePos>("Hidemaru_GetCursorPosUnicodeFromMousePos");
                    pGetCurrentWindowHandle = hmExeHandle.GetProcDelegate<TGetCurrentWindowHandle>("Hidemaru_GetCurrentWindowHandle");

                    if (Version >= 890)
                    {
                        pAnalyzeEncoding = hmExeHandle.GetProcDelegate<TAnalyzeEncoding>("Hidemaru_AnalyzeEncoding");
                        pLoadFileUnicode = hmExeHandle.GetProcDelegate<TLoadFileUnicode>("Hidemaru_LoadFileUnicode");
                    }
                    if (Version >= 915)
                    {
                        pGetStaticVariable = hmExeHandle.GetProcDelegate<TGetStaticVariable>("Hidemaru_GetStaticVariable");
                        pSetStaticVariable = hmExeHandle.GetProcDelegate<TSetStaticVariable>("Hidemaru_SetStaticVariable");
                    }

                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message);
                }

            }
        }
    }
}

namespace HmNetCOM
{
    namespace HmNativeMethods
    {
        internal partial class NativeMethods
        {
            [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
            protected extern static uint GetModuleFileName(IntPtr hModule, StringBuilder lpFilename, int nSize);

            [DllImport("kernel32.dll", SetLastError = true)]
            protected extern static IntPtr GlobalLock(IntPtr hMem);

            [DllImport("kernel32.dll", SetLastError = true)]

            [return: MarshalAs(UnmanagedType.Bool)]
            protected extern static bool GlobalUnlock(IntPtr hMem);

            [DllImport("kernel32.dll", SetLastError = true)]
            protected extern static IntPtr GlobalFree(IntPtr hMem);

            [StructLayout(LayoutKind.Sequential)]
            protected struct POINT
            {
                public int X;
                public int Y;
            }
            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            protected extern static bool GetCursorPos(out POINT lpPoint);

            [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true)]
            protected static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

            [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true, CharSet = CharSet.Unicode)]
            protected static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam);

            [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true, CharSet = CharSet.Unicode)]
            protected static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, StringBuilder wParam, StringBuilder lParam);
        }
    }

    internal partial class Hm : HmNativeMethods.NativeMethods
    {
    }
}

namespace HmNetCOM
{
    namespace HmNativeMethods
    {
        internal partial class NativeMethods
        {
            [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
            protected static extern IntPtr LoadLibrary(string lpFileName);

            [DllImport("kernel32", CharSet = CharSet.Ansi, BestFitMapping = false, ExactSpelling = true, SetLastError = true)]
            protected static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

            [DllImport("kernel32", SetLastError = true)]
            protected static extern bool FreeLibrary(IntPtr hModule);
        }
    }

    internal partial class Hm
    {
        // アンマネージドライブラリの遅延での読み込み。C++のLoadLibraryと同じことをするため
        // これをする理由は、この実行dllとHideamru.exeが異なるディレクトリに存在する可能性があるため、
        // C#風のDllImportは成立しないからだ。
        internal sealed class UnManagedDll : HmNativeMethods.NativeMethods, IDisposable
        {

            IntPtr moduleHandle;
            public UnManagedDll(string lpFileName)
            {
                moduleHandle = LoadLibrary(lpFileName);
            }

            // コード分析などの際の警告抑制のためにデストラクタをつけておく
            ~UnManagedDll()
            {
                // C#はメインドメインのdllは(このコードが存在するdll)はプロセスが終わらない限り解放されないので、
                // ここではネイティブのdllも事前には解放しない方がよい。(プロセス終了による解放に委ねる）
                // デストラクタでは何もしない。
                // コード分析でも警告がでないように、コード分析では実行されないことがわからない形で
                // 決して実行されないコードにしておく
                if (moduleHandle == (IntPtr)(-1)) { this.Dispose(); };
            }

            public IntPtr ModuleHandle
            {
                get
                {
                    return moduleHandle;
                }
            }

            public T GetProcDelegate<T>(string method) where T : class
            {
                IntPtr methodHandle = GetProcAddress(moduleHandle, method);
                T r = Marshal.GetDelegateForFunctionPointer(methodHandle, typeof(T)) as T;
                return r;
            }

            public void Dispose()
            {
                FreeLibrary(moduleHandle);
            }
        }

    }
}




/*
 * Copyright (C) 2021-2022 Akitsugu Komiyama
 * under the MIT License
 **/


namespace HmNetCOM
{
    internal partial class Hm
    {
        public static partial class Macro
        {
            /// <summary>
            /// マクロを実行中か否かを判定する
            /// </summary>
            /// <returns>実行中ならtrue, そうでなければfalse</returns>
            public static bool IsExecuting
            {
                get
                {
                    const int WM_USER = 0x400;
                    const int WM_ISMACROEXECUTING = WM_USER + 167;

                    IntPtr hWndHidemaru = WindowHandle;
                    if (hWndHidemaru != IntPtr.Zero)
                    {
                        IntPtr cwch = SendMessage(hWndHidemaru, WM_ISMACROEXECUTING, IntPtr.Zero, IntPtr.Zero);
                        return (cwch != IntPtr.Zero);
                    }

                    return false;
                }
            }

            /// <summary>
            /// マクロの静的な変数
            /// </summary>
            public static partial class StaticVar
            {

                /// <summary>
                /// マクロの静的な変数の値(文字列)を取得する
                /// </summary>
                /// <param name = "name">変数名</param>
                /// <param name = "sharedflag">共有フラグ</param>
                /// <returns>対象の静的変数名(name)に格納されている文字列</returns>
                public static string Get(string name, int sharedflag)
                {
                    return GetStaticVariable(name, sharedflag);
                }

                /// <summary>
                /// マクロの静的な変数へと値(文字列)を設定する
                /// </summary>
                /// <param name = "name">変数名</param>
                /// <param name = "value">設定する値(文字列)</param>
                /// <param name = "sharedflag">共有フラグ</param>
                /// <returns>取得に成功すれば真、失敗すれば偽が返る</returns>
                public static bool Set(string name, string value, int sharedflag)
                {
                    var ret = SetStaticVariable(name, value, sharedflag);
                    if (ret != 0)
                    {
                        return true;
                    }
                    return false;
                }

                private static int SetStaticVariable(String symbolname, String value, int sharedMemoryFlag)
                {
                    try
                    {
                        if (Version < 915)
                        {
                            throw new MissingMethodException("Hidemaru_Macro_SetGlobalVariable_Exception");
                        }
                        if (pSetStaticVariable == null)
                        {
                            throw new MissingMethodException("Hidemaru_Macro_SetGlobalVariable_Exception");
                        }

                        return pSetStaticVariable(symbolname, value, sharedMemoryFlag);
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Trace.WriteLine(e.Message);
                        throw;
                    }
                }

                private static string GetStaticVariable(String symbolname, int sharedMemoryFlag)
                {
                    try
                    {
                        if (Version < 915)
                        {
                            throw new MissingMethodException("Hidemaru_Macro_GetStaticVariable_Exception");
                        }
                        if (pGetStaticVariable == null)
                        {
                            throw new MissingMethodException("Hidemaru_Macro_GetStaticVariable_Exception");
                        }

                        string staticText = "";

                        IntPtr hGlobal = pGetStaticVariable(symbolname, sharedMemoryFlag);
                        if (hGlobal == IntPtr.Zero)
                        {
                            new InvalidOperationException("Hidemaru_Macro_GetStaticVariable_Exception");
                        }

                        var pwsz = GlobalLock(hGlobal);
                        if (pwsz != IntPtr.Zero)
                        {
                            staticText = Marshal.PtrToStringUni(pwsz);
                            GlobalUnlock(hGlobal);
                        }
                        GlobalFree(hGlobal);

                        return staticText;
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Trace.WriteLine(e.Message);
                        throw;
                    }
                }
            }

            /// <summary>
            /// マクロをプログラム内から実行した際の返り値のインターフェイス
            /// </summary>
            /// <returns>(Result, Message, Error)</returns>
            public interface IResult
            {
                int Result { get; }
                String Message { get; }
                Exception Error { get; }
            }

            private class TResult : IResult
            {
                public int Result { get; set; }
                public string Message { get; set; }
                public Exception Error { get; set; }

                public TResult(int Result, String Message, Exception Error)
                {
                    this.Result = Result;
                    this.Message = Message;
                    this.Error = Error;
                }
            }

            /// <summary>
            /// 現在のマクロ実行中に、プログラム中で、マクロを文字列で実行。
            /// マクロ実行中のみ実行可能なメソッド。
            /// </summary>
            /// <returns>(Result, Message, Error)</returns>

            public static IResult Eval(String expression)
            {
                TResult result;
                if (!IsExecuting)
                {
                    Exception e = new InvalidOperationException("Hidemaru_Macro_IsNotExecuting_Exception");
                    result = new TResult(-1, "", e);
                    return result;
                }
                int success = 0;
                try
                {
                    success = pEvalMacro(expression);
                }
                catch (Exception)
                {
                    throw;
                }
                if (success == 0)
                {
                    Exception e = new InvalidOperationException("Hidemaru_Macro_Eval_Exception");
                    result = new TResult(0, "", e);
                    return result;
                }
                else
                {
                    result = new TResult(success, "", null);
                    return result;
                }

            }

            public static partial class Exec
            {
                /// <summary>
                /// マクロを実行していない時に、プログラム中で、マクロファイルを与えて新たなマクロを実行。
                /// マクロを実行していない時のみ実行可能なメソッド。
                /// </summary>
                /// <returns>(Result, Message, Error)</returns>
                public static IResult File(string filepath)
                {
                    TResult result;
                    if (IsExecuting)
                    {
                        Exception e = new InvalidOperationException("Hidemaru_Macro_IsExecuting_Exception");
                        result = new TResult(-1, "", e);
                        return result;
                    }
                    if (!System.IO.File.Exists(filepath))
                    {
                        Exception e = new FileNotFoundException(filepath);
                        result = new TResult(-1, "", e);
                        return result;
                    }

                    const int WM_USER = 0x400;
                    const int WM_REMOTE_EXECMACRO_FILE = WM_USER + 271;
                    IntPtr hWndHidemaru = WindowHandle;

                    StringBuilder sbFileName = new StringBuilder(filepath);
                    StringBuilder sbRet = new StringBuilder("\x0f0f", 0x0f0f + 1); // 最初の値は帰り値のバッファー
                    IntPtr cwch = SendMessage(hWndHidemaru, WM_REMOTE_EXECMACRO_FILE, sbRet, sbFileName);
                    if (cwch != IntPtr.Zero)
                    {
                        result = new TResult(1, sbRet.ToString(), null);
                    }
                    else
                    {
                        Exception e = new InvalidOperationException("Hidemaru_Macro_Eval_Exception");
                        result = new TResult(0, sbRet.ToString(), e);
                    }
                    return result;
                }

                /// <summary>
                /// マクロを実行していない時に、プログラム中で、文字列で新たなマクロを実行。
                /// マクロを実行していない時のみ実行可能なメソッド。
                /// </summary>
                /// <returns>(Result, Message, Error)</returns>
                public static IResult Eval(string expression)
                {
                    TResult result;
                    if (IsExecuting)
                    {
                        Exception e = new InvalidOperationException("Hidemaru_Macro_IsExecuting_Exception");
                        result = new TResult(-1, "", e);
                        return result;
                    }

                    const int WM_USER = 0x400;
                    const int WM_REMOTE_EXECMACRO_MEMORY = WM_USER + 272;
                    IntPtr hWndHidemaru = WindowHandle;

                    StringBuilder sbExpression = new StringBuilder(expression);
                    StringBuilder sbRet = new StringBuilder("\x0f0f", 0x0f0f + 1); // 最初の値は帰り値のバッファー
                    IntPtr cwch = SendMessage(hWndHidemaru, WM_REMOTE_EXECMACRO_MEMORY, sbRet, sbExpression);
                    if (cwch != IntPtr.Zero)
                    {
                        result = new TResult(1, sbRet.ToString(), null);
                    }
                    else
                    {
                        Exception e = new InvalidOperationException("Hidemaru_Macro_Eval_Exception");
                        result = new TResult(0, sbRet.ToString(), e);
                    }
                    return result;
                }
            }
        }
    }
}

namespace HmNetCOM
{

    internal static class HmMacroExtentensions
    {
        public static void Deconstruct(this Hm.Macro.IResult result, out int Result, out Exception Error, out String Message)
        {
            Result = result.Result;
            Error = result.Error;
            Message = result.Message;
        }

        public static void Deconstruct(this Hm.Macro.IFunctionResult result, out object Result, out List<Object> Args, out Exception Error, out String Message)
        {
            Result = result.Result;
            Args = result.Args;
            Error = result.Error;
            Message = result.Message;
        }

        public static void Deconstruct(this Hm.Macro.IStatementResult result, out int Result, out List<Object> Args, out Exception Error, out String Message)
        {
            Result = result.Result;
            Args = result.Args;
            Error = result.Error;
            Message = result.Message;
        }
    }
}



/*
 * Copyright (C) 2021-2022 Akitsugu Komiyama
 * under the MIT License
 **/


namespace HmNetCOM
{
    internal partial class Hm
    {
        public static partial class Edit
        {
            /// <summary>
            /// キー入力があるなどの理由で処理を中断するべきかを返す。
            /// </summary>
            /// <returns>中断するべきならtrue、そうでなければfalse</returns>
            public static bool QueueStatus
            {
                get { return pCheckQueueStatus() != 0; }
            }

            /// <summary>
            /// 現在アクティブな編集領域のテキスト全体を返す。
            /// </summary>
            /// <returns>編集領域のテキスト全体</returns>
            public static string TotalText
            {
                get
                {
                    string totalText = "";
                    try
                    {
                        IntPtr hGlobal = pGetTotalTextUnicode();
                        if (hGlobal == IntPtr.Zero)
                        {
                            new InvalidOperationException("Hidemaru_GetTotalTextUnicode_Exception");
                        }

                        var pwsz = GlobalLock(hGlobal);
                        if (pwsz != IntPtr.Zero)
                        {
                            totalText = Marshal.PtrToStringUni(pwsz);
                            GlobalUnlock(hGlobal);
                        }
                        GlobalFree(hGlobal);
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                    return totalText;
                }
                set
                {
                    SetTotalText(value);
                }
            }
            static partial void SetTotalText(string text);


            /// <summary>
            /// 現在、単純選択している場合、その選択中のテキスト内容を返す。
            /// </summary>
            /// <returns>選択中のテキスト内容</returns>
            public static string SelectedText
            {
                get
                {
                    string selectedText = "";
                    try
                    {
                        IntPtr hGlobal = pGetSelectedTextUnicode();
                        if (hGlobal == IntPtr.Zero)
                        {
                            new InvalidOperationException("Hidemaru_GetSelectedTextUnicode_Exception");
                        }

                        var pwsz = GlobalLock(hGlobal);
                        if (pwsz != IntPtr.Zero)
                        {
                            selectedText = Marshal.PtrToStringUni(pwsz);
                            GlobalUnlock(hGlobal);
                        }
                        GlobalFree(hGlobal);
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                    return selectedText;
                }
                set
                {
                    SetSelectedText(value);
                }
            }
            static partial void SetSelectedText(string text);

            /// <summary>
            /// 現在、カーソルがある行(エディタ的)のテキスト内容を返す。
            /// </summary>
            /// <returns>カーソルがある行のテキスト内容</returns>
            public static string LineText
            {
                get
                {
                    string lineText = "";

                    ICursorPos pos = CursorPos;
                    if (pos.LineNo < 0 || pos.Column < 0)
                    {
                        return lineText;
                    }

                    try
                    {
                        IntPtr hGlobal = pGetLineTextUnicode(pos.LineNo);
                        if (hGlobal == IntPtr.Zero)
                        {
                            new InvalidOperationException("Hidemaru_GetLineTextUnicode_Exception");
                        }

                        var pwsz = GlobalLock(hGlobal);
                        if (pwsz != IntPtr.Zero)
                        {
                            lineText = Marshal.PtrToStringUni(pwsz);
                            GlobalUnlock(hGlobal);
                        }
                        GlobalFree(hGlobal);
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                    return lineText;
                }
                set
                {
                    SetLineText(value);
                }
            }
            static partial void SetLineText(string text);

            /// <summary>
            /// CursorPos の返り値のインターフェイス
            /// </summary>
            /// <returns>(LineNo, Column)</returns>
            public interface ICursorPos
            {
                int LineNo { get; }
                int Column { get; }
            }

            private struct TCursurPos : ICursorPos
            {
                public int Column { get; set; }
                public int LineNo { get; set; }
            }

            /// <summary>
            /// MousePos の返り値のインターフェイス
            /// </summary>
            /// <returns>(LineNo, Column, X, Y)</returns>
            public interface IMousePos
            {
                int LineNo { get; }
                int Column { get; }
                int X { get; }
                int Y { get; }
            }

            private struct TMousePos : IMousePos
            {
                public int LineNo { get; set; }
                public int Column { get; set; }
                public int X { get; set; }
                public int Y { get; set; }
            }

            /// <summary>
            /// ユニコードのエディタ的な換算でのカーソルの位置を返す
            /// </summary>
            /// <returns>(LineNo, Column)</returns>
            public static ICursorPos CursorPos
            {
                get
                {
                    int lineno = -1;
                    int column = -1;
                    int success = pGetCursorPosUnicode(out lineno, out column);
                    if (success != 0)
                    {
                        TCursurPos pos = new TCursurPos();
                        pos.LineNo = lineno;
                        pos.Column = column;
                        return pos;
                    }
                    else
                    {
                        TCursurPos pos = new TCursurPos();
                        pos.LineNo = -1;
                        pos.Column = -1;
                        return pos;
                    }

                }
            }

            /// <summary>
            /// ユニコードのエディタ的な換算でのマウスの位置に対応するカーソルの位置を返す
            /// </summary>
            /// <returns>(LineNo, Column, X, Y)</returns>
            public static IMousePos MousePos
            {
                get
                {
                    POINT lpPoint;
                    bool success_1 = GetCursorPos(out lpPoint);

                    TMousePos pos = new TMousePos
                    {
                        LineNo = -1,
                        Column = -1,
                        X = -1,
                        Y = -1,
                    };

                    if (!success_1)
                    {
                        return pos;
                    }

                    int column = -1;
                    int lineno = -1;
                    int success_2 = pGetCursorPosUnicodeFromMousePos(IntPtr.Zero, out lineno, out column);
                    if (success_2 == 0)
                    {
                        return pos;
                    }

                    pos.LineNo = lineno;
                    pos.Column = column;
                    pos.X = lpPoint.X;
                    pos.Y = lpPoint.Y;
                    return pos;
                }
            }

            /// <summary>
            /// 現在開いているファイル名のフルパスを返す、無題テキストであれば、nullを返す。
            /// </summary>
            /// <returns>ファイル名のフルパス、もしくは null</returns>

            public static string FilePath
            {
                get
                {
                    IntPtr hWndHidemaru = WindowHandle;
                    if (hWndHidemaru != IntPtr.Zero)
                    {
                        const int WM_USER = 0x400;
                        const int WM_HIDEMARUINFO = WM_USER + 181;
                        const int HIDEMARUINFO_GETFILEFULLPATH = 4;

                        StringBuilder sb = new StringBuilder(filePathMaxLength); // まぁこんくらいでさすがに十分なんじゃないの...
                        IntPtr cwch = SendMessage(hWndHidemaru, WM_HIDEMARUINFO, new IntPtr(HIDEMARUINFO_GETFILEFULLPATH), sb);
                        String filename = sb.ToString();
                        if (String.IsNullOrEmpty(filename))
                        {
                            return null;
                        }
                        else
                        {
                            return filename;
                        }
                    }
                    return null;
                }
            }

            /// <summary>
            /// 現在開いている編集エリアで、文字列の編集や何らかの具体的操作を行ったかチェックする。マクロ変数のupdatecount相当
            /// </summary>
            /// <returns>一回の操作でも数カウント上がる。32bitの値を超えると一周する。初期値は1以上。</returns>

            public static int UpdateCount
            {
                get
                {
                    if (Version < 912.98)
                    {
                        throw new MissingMethodException("Hidemaru_Edit_UpdateCount_Exception");
                    }
                    IntPtr hWndHidemaru = WindowHandle;
                    if (hWndHidemaru != IntPtr.Zero)
                    {
                        const int WM_USER = 0x400;
                        const int WM_HIDEMARUINFO = WM_USER + 181;
                        const int HIDEMARUINFO_GETUPDATECOUNT = 7;

                        IntPtr updatecount = SendMessage(hWndHidemaru, WM_HIDEMARUINFO, (IntPtr)HIDEMARUINFO_GETUPDATECOUNT, IntPtr.Zero);
                        return (int)updatecount;
                    }
                    return -1;
                }
            }

        }
    }
}


namespace HmNetCOM
{
    internal partial class Hm
    {
        public static partial class Edit
        {
            static partial void SetTotalText(string text)
            {
                string myDllFullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string myTargetDllFullPath = HmMacroCOMVar.GetMyTargetDllFullPath(myDllFullPath);
                string myTargetClass = HmMacroCOMVar.GetMyTargetClass(myDllFullPath);
                HmMacroCOMVar.SetMacroVar(text);
                string cmd = $@"
                begingroupundo;
                selectall;
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                insert member(#_COM_NET_PINVOKE_MACRO_VAR, ""DllToMacro"" );
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
                endgroupundo;
                ";
                Macro.IResult result = null;
                if (Macro.IsExecuting)
                {
                    result = Hm.Macro.Eval(cmd);
                }
                else
                {
                    result = Hm.Macro.Exec.Eval(cmd);
                }

                HmMacroCOMVar.ClearVar();
                if (result.Error != null)
                {
                    throw result.Error;
                }
            }

            static partial void SetSelectedText(string text)
            {
                string myDllFullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string myTargetDllFullPath = HmMacroCOMVar.GetMyTargetDllFullPath(myDllFullPath);
                string myTargetClass = HmMacroCOMVar.GetMyTargetClass(myDllFullPath);
                HmMacroCOMVar.SetMacroVar(text);
                string cmd = $@"
                if (selecting) {{
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                insert member(#_COM_NET_PINVOKE_MACRO_VAR, ""DllToMacro"" );
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
                }}
                ";

                Macro.IResult result = null;
                if (Macro.IsExecuting)
                {
                    result = Hm.Macro.Eval(cmd);
                }
                else
                {
                    result = Hm.Macro.Exec.Eval(cmd);
                }

                HmMacroCOMVar.ClearVar();
                if (result.Error != null)
                {
                    throw result.Error;
                }
            }

            static partial void SetLineText(string text)
            {
                string myDllFullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string myTargetDllFullPath = HmMacroCOMVar.GetMyTargetDllFullPath(myDllFullPath);
                string myTargetClass = HmMacroCOMVar.GetMyTargetClass(myDllFullPath);
                HmMacroCOMVar.SetMacroVar(text);
                var pos = Edit.CursorPos;
                string cmd = $@"
                begingroupundo;
                selectline;
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                insert member(#_COM_NET_PINVOKE_MACRO_VAR, ""DllToMacro"" );
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
                moveto2 {pos.Column}, {pos.LineNo};
                endgroupundo;
                ";

                Macro.IResult result = null;
                if (Macro.IsExecuting)
                {
                    result = Hm.Macro.Eval(cmd);
                }
                else
                {
                    result = Hm.Macro.Exec.Eval(cmd);
                }

                HmMacroCOMVar.ClearVar();
                if (result.Error != null)
                {
                    throw result.Error;
                }
            }

        }
    }
}


namespace HmNetCOM
{

    internal static class HmEditExtentensions
    {
        public static void Deconstruct(this Hm.Edit.ICursorPos pos, out int LineNo, out int Column)
        {
            LineNo = pos.LineNo;
            Column = pos.Column;
        }

        public static void Deconstruct(this Hm.Edit.IMousePos pos, out int LineNo, out int Column, out int X, out int Y)
        {
            LineNo = pos.LineNo;
            Column = pos.Column;
            X = pos.X;
            Y = pos.Y;
        }
    }
}



/*
 * Copyright (C) 2021-2022 Akitsugu Komiyama
 * under the MIT License
 **/



namespace HmNetCOM
{
    internal partial class Hm
    {
        public static partial class File
        {
            public interface IHidemaruEncoding
            {
                int HmEncode { get; }
            }
            public interface IMicrosoftEncoding
            {
                int MsCodePage { get; }
            }

            public interface IEncoding : IHidemaruEncoding, IMicrosoftEncoding
            {
            }

            public interface IHidemaruStreamReader : IDisposable
            {
                IEncoding Encoding { get; }
                String Read();
                String FilePath { get; }
                void Close();
            }

            // 途中でエラーが出るかもしれないので、相応しいUnlockやFreeが出来るように内部管理用
            private enum HGlobalStatus { None, Lock, Unlock, Free };

            private static String ReadAllText(String filepath, int hm_encode = -1)
            {
                if (pLoadFileUnicode == null)
                {
                    throw new MissingMethodException("Hidemaru_LoadFileUnicode");
                }

                if (hm_encode == -1)
                {
                    hm_encode = GetHmEncode(filepath);
                }

                if (!System.IO.File.Exists(filepath))
                {
                    throw new System.IO.FileNotFoundException(filepath);
                }

                String curstr = "";
                int read_count = 0;
                IntPtr hGlobal = pLoadFileUnicode(filepath, hm_encode, ref read_count, IntPtr.Zero, IntPtr.Zero);
                HGlobalStatus hgs = HGlobalStatus.None;
                if (hGlobal == IntPtr.Zero)
                {
                    throw new System.IO.IOException(filepath);
                }
                if (hGlobal != IntPtr.Zero)
                {
                    try
                    {
                        IntPtr ret = GlobalLock(hGlobal);
                        hgs = HGlobalStatus.Lock;
                        curstr = Marshal.PtrToStringUni(ret);
                        GlobalUnlock(hGlobal);
                        hgs = HGlobalStatus.Unlock;
                        GlobalFree(hGlobal);
                        hgs = HGlobalStatus.Free;
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Trace.WriteLine(e.Message);
                    }
                    finally
                    {
                        switch (hgs)
                        {
                            // ロックだけ成功した
                            case HGlobalStatus.Lock:
                                {
                                    GlobalUnlock(hGlobal);
                                    GlobalFree(hGlobal);
                                    break;
                                }
                            // アンロックまで成功した
                            case HGlobalStatus.Unlock:
                                {
                                    GlobalFree(hGlobal);
                                    break;
                                }
                            // フリーまで成功した
                            case HGlobalStatus.Free:
                                {
                                    break;
                                }
                        }
                    }
                }
                if (hgs == HGlobalStatus.Free)
                {
                    return curstr;
                }
                else
                {
                    throw new System.IO.IOException(filepath);
                }
            }

            private static int[] key_encode_value_codepage_array = {
                0,      // Unknown
                932,    // encode = 1 ANSI/OEM Japanese; Japanese (Shift-JIS)
                1200,   // encode = 2 Unicode UTF-16, little-endian
                51932,  // encode = 3 EUC
                50221,  // encode = 4 JIS
                65000,  // encode = 5 UTF-7
                65001,  // encode = 6 UTF-8
                1201,   // encode = 7 Unicode UTF-16, big-endian
                1252,   // encode = 8 欧文 ANSI Latin 1; Western European (Windows)
                936,    // encode = 9 簡体字中国語 ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GB2312)
                950,    // encode =10 繁体字中国語 ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)
                949,    // encode =11 韓国語 ANSI/OEM Korean (Unified Hangul Code)
                1361,   // encode =12 韓国語 Korean (Johab)
                1250,   // encode =13 中央ヨーロッパ言語 ANSI Central European; Central European (Windows)
                1257,   // encode =14 バルト語 ANSI Baltic; Baltic (Windows)
                1253,   // encode =15 ギリシャ語 ANSI Greek; Greek (Windows)
                1251,   // encode =16 キリル言語 ANSI Cyrillic; Cyrillic (Windows)
                42,     // encode =17 シンボル
                1254,   // encode =18 トルコ語 ANSI Turkish; Turkish (Windows)
                1255,   // encode =19 ヘブライ語 ANSI Hebrew; Hebrew (Windows)
                1256,   // encode =20 アラビア語 ANSI Arabic; Arabic (Windows)
                874,    // encode =21 タイ語 ANSI/OEM Thai (same as 28605, ISO 8859-15); Thai (Windows)
                1258,   // encode =22 ベトナム語 ANSI/OEM Vietnamese; Vietnamese (Windows)
                10001,  // encode =23 x-mac-japanese Japanese (Mac)
                850,    // encode =24 OEM/DOS
                0,      // encode =25 その他
                12000,  // encode =26 Unicode (UTF-32) little-endian
                12001,  // encode =27 Unicode (UTF-32) big-endian

            };

            /// <summary>
            /// 秀丸でファイルのエンコードを取得する
            /// （秀丸に設定されている内容に従う）
            /// </summary>
            /// <returns>IEncoding型のオブジェクト。MsCodePage と HmEncode のプロパティを得られる。</returns>
            public static IEncoding GetEncoding(string filepath)
            {
                int hm_encode = GetHmEncode(filepath);
                int ms_codepage = GetMsCodePage(hm_encode);
                IEncoding encoding = new Encoding(hm_encode, ms_codepage);
                return encoding;
            }

            private static int GetHmEncode(string filepath)
            {

                if (pAnalyzeEncoding == null)
                {
                    throw new MissingMethodException("Hidemaru_AnalyzeEncoding");
                }

                return pAnalyzeEncoding(filepath, IntPtr.Zero, IntPtr.Zero);
            }

            private static int GetMsCodePage(int hidemaru_encode)
            {
                int result_codepage = 0;

                if (pAnalyzeEncoding == null)
                {
                    throw new MissingMethodException("Hidemaru_AnalyzeEncoding");
                }

                /*
                 *
                    Shift-JIS encode=1 codepage=932
                    Unicode encode=2 codepage=1200
                    EUC encode=3 codepage=51932
                    JIS encode=4 codepage=50221
                    UTF-7 encode=5 codepage=65000
                    UTF-8 encode=6 codepage=65001
                    Unicode (Big-Endian) encode=7 codepage=1201
                    欧文 encode=8 codepage=1252
                    簡体字中国語 encode=9 codepage=936
                    繁体字中国語 encode=10 codepage=950
                    韓国語 encode=11 codepage=949
                    韓国語(Johab) encode=12 codepage=1361
                    中央ヨーロッパ言語 encode=13 codepage=1250
                    バルト語 encode=14 codepage=1257
                    ギリシャ語 encode=15 codepage=1253
                    キリル言語 encode=16 codepage=1251
                    シンボル encode=17 codepage=42
                    トルコ語 encode=18 codepage=1254
                    ヘブライ語 encode=19 codepage=1255
                    アラビア語 encode=20 codepage=1256
                    タイ語 encode=21 codepage=874
                    ベトナム語 encode=22 codepage=1258
                    Macintosh encode=23 codepage=0
                    OEM/DOS encode=24 codepage=0
                    その他 encode=25 codepage=0
                    UTF-32 encode=27 codepage=12000
                    UTF-32 (Big-Endian) encode=28 codepage=12001
                */
                if (hidemaru_encode <= 0)
                {
                    return result_codepage;
                }

                if (hidemaru_encode < key_encode_value_codepage_array.Length)
                {
                    // 把握しているコードページなので入れておく
                    result_codepage = key_encode_value_codepage_array[hidemaru_encode];
                    return result_codepage;
                }
                else // 長さ以上なら、予期せぬ未来のencode番号対応
                {
                    return result_codepage;
                }

            }

            // コードページを得る
            private static int GetMsCodePage(string filepath)
            {

                int result_codepage = 0;

                if (pAnalyzeEncoding == null)
                {
                    throw new MissingMethodException("Hidemaru_AnalyzeEncoding");
                }


                int hidemaru_encode = pAnalyzeEncoding(filepath, IntPtr.Zero, IntPtr.Zero);

                /*
                 *
                    Shift-JIS encode=1 codepage=932
                    Unicode encode=2 codepage=1200
                    EUC encode=3 codepage=51932
                    JIS encode=4 codepage=50221
                    UTF-7 encode=5 codepage=65000
                    UTF-8 encode=6 codepage=65001
                    Unicode (Big-Endian) encode=7 codepage=1201
                    欧文 encode=8 codepage=1252
                    簡体字中国語 encode=9 codepage=936
                    繁体字中国語 encode=10 codepage=950
                    韓国語 encode=11 codepage=949
                    韓国語(Johab) encode=12 codepage=1361
                    中央ヨーロッパ言語 encode=13 codepage=1250
                    バルト語 encode=14 codepage=1257
                    ギリシャ語 encode=15 codepage=1253
                    キリル言語 encode=16 codepage=1251
                    シンボル encode=17 codepage=42
                    トルコ語 encode=18 codepage=1254
                    ヘブライ語 encode=19 codepage=1255
                    アラビア語 encode=20 codepage=1256
                    タイ語 encode=21 codepage=874
                    ベトナム語 encode=22 codepage=1258
                    Macintosh encode=23 codepage=0
                    OEM/DOS encode=24 codepage=0
                    その他 encode=25 codepage=0
                    UTF-32 encode=27 codepage=12000
                    UTF-32 (Big-Endian) encode=28 codepage=12001
                */
                if (hidemaru_encode <= 0)
                {
                    return result_codepage;
                }

                if (hidemaru_encode < key_encode_value_codepage_array.Length)
                {
                    // 把握しているコードページなので入れておく
                    result_codepage = key_encode_value_codepage_array[hidemaru_encode];
                    return result_codepage;
                }
                else // 長さ以上なら、予期せぬ未来のencode番号対応
                {
                    return result_codepage;
                }
            }

            private class Encoding : IEncoding
            {
                private int m_hm_encode;
                private int m_ms_codepage;
                public Encoding(int hmencode, int mscodepage)
                {
                    this.m_hm_encode = hmencode;
                    this.m_ms_codepage = mscodepage;
                }
                public int HmEncode { get { return this.m_hm_encode; } }
                public int MsCodePage { get { return this.m_ms_codepage; } }
            }

            private class HidemaruStreamReader : IHidemaruStreamReader
            {
                String m_path;

                IEncoding m_encoding;

                Hm.File.IEncoding Hm.File.IHidemaruStreamReader.Encoding { get { return this.m_encoding; } }

                public string FilePath { get { return this.m_path; } }

                public HidemaruStreamReader(String path, int hm_encode = -1)
                {
                    this.m_path = path;
                    // 指定されていなければ、
                    if (hm_encode == -1)
                    {
                        hm_encode = GetHmEncode(path);
                    }
                    int ms_codepage = GetMsCodePage(hm_encode);
                    this.m_encoding = new Encoding(hm_encode, ms_codepage);
                }

                ~HidemaruStreamReader()
                {
                    Close();
                }

                String IHidemaruStreamReader.Read()
                {
                    if (System.IO.File.Exists(this.m_path) == false)
                    {
                        throw new System.IO.FileNotFoundException(this.m_path);
                    }

                    try
                    {
                        String text = ReadAllText(this.m_path, this.m_encoding.HmEncode);
                        return text;
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }

                public void Close()
                {
                    if (this.m_path != null)
                    {
                        this.m_encoding = null;
                        this.m_path = null;
                    }
                }

                public void Dispose()
                {
                    this.Close();
                }
            }

            /// <summary>
            /// 秀丸でファイルのエンコードを判断し、その判断結果に基づいてファイルのテキスト内容を取得する。
            /// （秀丸に設定されている内容に従う）
            /// </summary>
            /// <param name = "filepath">読み込み対象のファイルのパス</param>
            /// <param name = "hm_encode">エンコード(秀丸マクロの「encode」の値)が分かっているならば指定する、指定しない場合秀丸APIの自動判定に任せる。</param>
            /// <returns>IHidemaruStreamReader型のオブジェクト。</returns>
            public static IHidemaruStreamReader Open(string filepath, int hm_encode = -1)
            {
                if (System.IO.File.Exists(filepath) == false)
                {
                    throw new System.IO.FileNotFoundException(filepath);
                }
                var sr = new HidemaruStreamReader(filepath, hm_encode);
                return sr;
            }

        }
    }
}




/*
 * Copyright (C) 2021-2022 Akitsugu Komiyama
 * under the MIT License
 **/


namespace HmNetCOM
{
    // このインターフェイスは秀丸マクロのjsmode(WebView2)でCOMを呼び出す際に必要
    interface IHmMacroCOMVar
    {
        object DllToMacro();
        int MacroToDll(object variable);
        int MethodToDll(String dllfullpath, String typefullname, String methodname, String message_param);
    }

    public partial class HmMacroCOMVar
    {
        private const string HmMacroCOMVarInterface = "28289750-2D0C-4EB5-86D4-630EE52BE000";
    }
}




namespace HmNetCOM
{
    // 秀丸のCOMから呼び出して、マクロ⇔COMといったように、マクロとプログラムで変数値を互いに伝搬する
    [ComVisible(true)]
#if (NET || NETCOREAPP3_1)
#else
    [ClassInterface(ClassInterfaceType.None)]
#endif
    [Guid(HmMacroCOMVarInterface)]
    public partial class HmMacroCOMVar : IHmMacroCOMVar, Hm.IComSupportX64
    {
        private static object marcroVar = null;
        public object DllToMacro()
        {
            return marcroVar;
        }
        public int MacroToDll(object variable)
        {
            marcroVar = variable;
            return 1;
        }
        public int MethodToDll(String dllfullpath, String typefullname, String methodname, String message_param)
        {
            marcroVar = message_param;

            try
            {
                MethodToDllHelper(dllfullpath, typefullname, methodname, message_param);
                return 1;
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine(e);
            }
            return 0;
        }

        private void TraceMethodInfo(String assm_path, String class_name, String method_name)
        {
            System.Diagnostics.Trace.WriteLine("アセンブリパス   :" + assm_path);
            System.Diagnostics.Trace.WriteLine("名前空間.クラス名:" + class_name);
            System.Diagnostics.Trace.WriteLine("メソッド名       :" + method_name);
        }
        private static void TraceExceptionInfo(Exception e)
        {
            System.Diagnostics.Trace.WriteLine(e.GetType());
            System.Diagnostics.Trace.WriteLine(e.Message);
            System.Diagnostics.Trace.WriteLine(e.StackTrace);
        }
        private Object MethodToDllHelper(String assm_path, String class_name, String method_name, String message_param)
        {
            Exception method_ex = null;
            try
            {
                Assembly assm = null;
                Type t = null;

                if (assm_path.Length > 0)
                {
                    assm = Assembly.LoadFile(assm_path);
                    if (assm == null)
                    {
                        System.Diagnostics.Trace.WriteLine("ロード出来ない");
                    }
                    else
                    {
                        // System::Diagnostics::Trace::WriteLine(assm->FullName);
                    }

                    foreach (Type t2 in assm.GetExportedTypes())
                    {
                        if (t2.ToString() == class_name)
                        {
                            t = assm.GetType(class_name);
                        }
                    }
                }
                else
                {
                    t = Type.GetType(class_name);
                }
                if (t == null)
                {
                    System.Diagnostics.Trace.WriteLine("MissingMethodException(クラスもしくはメソッドを見つけることが出来ない):");
                    TraceMethodInfo(assm_path, class_name, method_name);
                    return null;
                }

                // メソッドの定義タイプを探る。
                MethodInfo m;
                try
                {
                    m = t.GetMethod(method_name);
                }
                catch (Exception ex)
                {
                    // 基本コースだと一致してない系の可能性やオーバーロードなど未解決エラーを控えておく
                    // t->GetMethod(...)は論理的には不要だが、「エラー情報のときにわかりやすい情報を.NETに自動で出力してもらう」ためにダミーで呼び出しておく
                    method_ex = ex;

                    // オーバーロードなら1つに解決できるように型情報も含めてmは上書き
                    List<Type> args_types = new List<Type>();
                    args_types.Add(Type.GetType(message_param));
                    m = t.GetMethod(method_name, args_types.ToArray());
                }

                Object o = null;
                try
                {
                    // オーバーロードなら1つに解決できるように型情報も含めてmは上書き
                    List<Object> args_values = new List<Object>();
                    args_values.Add(message_param);
                    o = m.Invoke(null, args_values.ToArray());
                }
                catch (Exception)
                {
                    System.Diagnostics.Trace.WriteLine("指定のメソッドの実行時、例外が発生しました。");
                    throw;
                }
                return o;
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine("指定のアセンブリやメソッドを特定する前に、例外が発生しました。");
                TraceMethodInfo(assm_path, class_name, method_name);
                if (method_ex != null)
                {
                    TraceExceptionInfo(method_ex);
                }
                TraceExceptionInfo(e);
            }

            return null;

        }
        public bool X64MACRO()
        {
            return true;
        }
    }

    public partial class HmMacroCOMVar
    {
        static HmMacroCOMVar()
        {
            var h = new HmMacroCOMVar();
            myGuidLabel = h.GetType().GUID.ToString();
            myClassFullName = h.GetType().FullName;
        }

        internal static void SetMacroVar(object obj)
        {
            marcroVar = obj;
        }
        internal static object GetMacroVar()
        {
            return marcroVar;
        }
        private static string myGuidLabel = "";
        private static string myClassFullName = "";

        internal static string GetMyTargetDllFullPath(string thisDllFullPath)
        {
            string myTargetClass = myClassFullName;
            string thisComHostFullPath = System.IO.Path.ChangeExtension(thisDllFullPath, "comhost.dll");
            if (System.IO.File.Exists(thisComHostFullPath))
            {
                return thisComHostFullPath;
            }

            return thisDllFullPath;
        }

        internal static string GetMyTargetClass(string thisDllFullPath)
        {
            string myTargetClass = myClassFullName;
            string thisComHostFullPath = System.IO.Path.ChangeExtension(thisDllFullPath, "comhost.dll");
            if (System.IO.File.Exists(thisComHostFullPath))
            {
                myTargetClass = "{" + myGuidLabel + "}";
            }

            return myTargetClass;
        }

        internal static object GetVar(string var_name)
        {
            string myDllFullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string myTargetDllFullPath = GetMyTargetDllFullPath(myDllFullPath);
            string myTargetClass = GetMyTargetClass(myDllFullPath);
            ClearVar();
            var result = Hm.Macro.Eval($@"
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                #_COM_NET_PINVOKE_MACRO_VAR_RESULT = member(#_COM_NET_PINVOKE_MACRO_VAR, ""MacroToDll"", {var_name});
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
                #_COM_NET_PINVOKE_MACRO_VAR_RESULT = 0;
            ");
            if (result.Error != null)
            {
                throw result.Error;
            }
            return HmMacroCOMVar.marcroVar;
        }

        internal static int SetVar(string var_name, object obj)
        {
            string myDllFullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string myTargetDllFullPath = GetMyTargetDllFullPath(myDllFullPath);
            string myTargetClass = GetMyTargetClass(myDllFullPath);
            ClearVar();
            HmMacroCOMVar.marcroVar = obj;
            var result = Hm.Macro.Eval($@"
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                {var_name} = member(#_COM_NET_PINVOKE_MACRO_VAR, ""DllToMacro"" );
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
            ");
            if (result.Error != null)
            {
                throw result.Error;
            }
            return 1;
        }

        internal static void ClearVar()
        {
            HmMacroCOMVar.marcroVar = null;
        }

        internal static Hm.Macro.IResult BornMacroScopeMethod(String scopename, String dllfullpath, String typefullname, String methodname)
        {

            string myDllFullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string myTargetDllFullPath = GetMyTargetDllFullPath(myDllFullPath);
            string myTargetClass = GetMyTargetClass(myDllFullPath);
            ClearVar();
            var result = Hm.Macro.Exec.Eval($@"
                #_COM_NET_PINVOKE_METHOD_CALL = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                #_COM_NET_PINVOKE_METHOD_CALL_RESULT = member(#_COM_NET_PINVOKE_METHOD_CALL, ""MethodToDll"", @""{dllfullpath}"", @""{typefullname}"", @""{methodname}"",  R""MACRO_OF_SCOPENAME({scopename})MACRO_OF_SCOPENAME"");
                releaseobject(#_COM_NET_PINVOKE_METHOD_CALL);
                #_COM_NET_PINVOKE_METHOD_CALL_RESULT = 0;
            ");
            return result;
        }
    }


    internal partial class Hm
    {
        public static partial class Macro
        {

            public static partial class Exec
            {
                /// <summary>
                /// 指定のC#のstaticメソッドを「新たなマクロ実行空間」として呼び出す
                /// </summary>
                /// <param name = "message_parameter">文字列パラメータ</param>
                /// <param name = "delegate_method">呼び出したいC#メソッド「public methodname(string message_parameter)の型」に従うメソッドであること</param>
                /// <returns>(Result, Message, Error)</returns>
                public static IResult Method(string message_parameter, Delegate delegate_method)
                {
                    string parameter = message_parameter;
                    // 渡されたメソッドが自分自身のdllと異なるのはダメ
                    if (delegate_method.Method.DeclaringType.Assembly.Location != System.Reflection.Assembly.GetExecutingAssembly().Location)
                    {
                        string message_no_dll_myself = "The Delegate method must in " + System.Reflection.Assembly.GetExecutingAssembly().Location;
                        var result_no_dll_myself = new TResult(0, "", new MissingMethodException(message_no_dll_myself));
                        System.Diagnostics.Trace.WriteLine(result_no_dll_myself);
                        return result_no_dll_myself;
                    }
                    else if (delegate_method.Method.IsStatic && delegate_method.Method.IsPublic)
                    {
                        var ret = HmMacroCOMVar.BornMacroScopeMethod(parameter, delegate_method.Method.DeclaringType.Assembly.Location, delegate_method.Method.DeclaringType.FullName, delegate_method.Method.Name);
                        if (ret.Result > 0)
                        {
                            var result = new TResult(ret.Result, message_parameter, ret.Error);
                            return result;
                        }
                        else
                        {
                            var result = new TResult(ret.Result, ret.Message, ret.Error);
                            return result;
                        }
                    }
                    else if (!delegate_method.Method.IsStatic)
                    {

                        string message_no_static = delegate_method.Method.DeclaringType.FullName + "." + delegate_method.Method.Name + " is not 'STATIC' in " + delegate_method.Method.DeclaringType.Assembly.Location;
                        var result_no_static = new TResult(0, "", new MissingMethodException(message_no_static));
                        System.Diagnostics.Trace.WriteLine(message_no_static);
                        return result_no_static;
                    }
                    else if (!delegate_method.Method.IsPublic)
                    {
                        string message_no_public = delegate_method.Method.DeclaringType.FullName + "." + delegate_method.Method.Name + " is not 'PUBLIC' in " + delegate_method.Method.DeclaringType.Assembly.Location;
                        var result_no_public = new TResult(0, "", new MissingMethodException(message_no_public));
                        System.Diagnostics.Trace.WriteLine(message_no_public);
                        return result_no_public;
                    }
                    string message_missing = delegate_method.Method.DeclaringType.FullName + "." + delegate_method.Method.Name + "is 'MISSING' access in " + delegate_method.Method.DeclaringType.Assembly.Location;
                    var result_missing = new TResult(0, "", new MissingMethodException(delegate_method.Method.Name + " is missing access"));
                    System.Diagnostics.Trace.WriteLine(result_missing);
                    return result_missing;
                }
            }
        }


        public static partial class Macro
        {
            // マクロでの問い合わせ結果系
            public interface IStatementResult
            {
                int Result { get; }
                String Message { get; }
                Exception Error { get; }
                List<Object> Args { get; }
            }


            private class TStatementResult : IStatementResult
            {
                public int Result { get; set; }
                public string Message { get; set; }
                public Exception Error { get; set; }
                public List<Object> Args { get; set; }

                public TStatementResult(int Result, String Message, Exception Error, List<Object> Args)
                {
                    this.Result = Result;
                    this.Message = Message;
                    this.Error = Error;
                    this.Args = new List<object>(Args); // コピー渡し
                }
            }

            private static int statement_base_random = 0;
            /// <summary>
            /// 秀丸マクロの関数のような「命令文」を実行
            /// </summary>
            /// <param name = "statement_name">（関数のような）命令文名</param>
            /// <param name = "args">命令文の引数</param>
            /// <returns>(Result, Args, Message, Error)</returns>
            internal static IStatementResult Statement(string statement_name, params object[] args)
            {
                string funcname = statement_name;
                if (statement_base_random == 0)
                {
                    statement_base_random = new System.Random().Next(Int16.MaxValue) + 1;

                }

                List<KeyValuePair<string, object>> arg_list = SetMacroVarAndMakeMacroKeyArray(args, statement_base_random);

                // keyをリスト化
                var arg_keys = new List<String>();
                foreach (var l in arg_list)
                {
                    arg_keys.Add(l.Key);
                }

                // それを「,」で繋げる
                string args_string = String.Join(", ", arg_keys);
                // それを指定の「文」で実行する形
                string expression = $"{funcname} {args_string};\n";

                // 実行する
                IResult ret = Macro.Eval(expression);

                int macro_result = ret.Result;
                if (ret.Error == null)
                {
                    try
                    {
                        Object tmp_var = Macro.Var["result"]; // この中のGetMethodで例外が発生する可能性あり

                        if (IntPtr.Size == 4)
                        {
                            macro_result = (Int32)tmp_var + 0; // 確実に複製を
                        }
                        else
                        {
                            Int64 macro_result64 = (Int64)tmp_var + 0; // 確実に複製を
                            Int32 macro_result32 = (Int32)HmClamp<Int64>(macro_result64, Int32.MinValue, Int32.MaxValue);
                            macro_result = (Int32)macro_result32;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                // 成否も含めて結果を入れる。
                IStatementResult result = new TStatementResult(macro_result, ret.Message, ret.Error, new List<Object>());

                // 使ったので削除
                for (int ix = 0; ix < arg_list.Count; ix++)
                {
                    var l = arg_list[ix];
                    if (l.Value is Int32 || l.Value is Int64)
                    {
                        result.Args.Add(Macro.Var[l.Key]);
                        Macro.Var[l.Key] = 0;
                    }
                    else if (l.Value is string)
                    {
                        result.Args.Add(Macro.Var[l.Key]);
                        Macro.Var[l.Key] = "";
                    }

                    else if (l.Value.GetType() == new List<int>().GetType() || l.Value.GetType() == new List<long>().GetType() || l.Value.GetType() == new List<IntPtr>().GetType())
                    {
                        result.Args.Add(l.Value);
                        if (l.Value.GetType() == new List<int>().GetType())
                        {
                            List<int> int_list = (List<int>)l.Value;
                            for (int iix = 0; iix < int_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                        else if (l.Value.GetType() == new List<long>().GetType())
                        {
                            List<long> long_list = (List<long>)l.Value;
                            for (int iix = 0; iix < long_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                        else if (l.Value.GetType() == new List<IntPtr>().GetType())
                        {
                            List<IntPtr> ptr_list = (List<IntPtr>)l.Value;
                            for (int iix = 0; iix < ptr_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                    }
                    else if (l.Value.GetType() == new List<String>().GetType())
                    {
                        result.Args.Add(l.Value);
                        List<String> ptr_list = (List<String>)l.Value;
                        for (int iix = 0; iix < ptr_list.Count; iix++)
                        {
                            Macro.Var[l.Key + "[" + iix + "]"] = "";
                        }
                    }
                    else
                    {
                        result.Args.Add(l.Value);
                    }
                }

                return result;
            }

            // マクロでの問い合わせ結果系
            public interface IFunctionResult
            {
                object Result { get; }
                String Message { get; }
                Exception Error { get; }
                List<Object> Args { get; }
            }

            private class TFunctionResult : IFunctionResult
            {
                public object Result { get; set; }
                public string Message { get; set; }
                public Exception Error { get; set; }
                public List<Object> Args { get; set; }

                public TFunctionResult(object Result, String Message, Exception Error, List<Object> Args)
                {
                    this.Result = Result;
                    this.Message = Message;
                    this.Error = Error;
                    this.Args = new List<object>(Args); // コピー渡し
                }
            }

            private static int funciton_base_random = 0;
            /// <summary>
            /// 秀丸マクロの「関数」を実行
            /// </summary>
            /// <param name = "func_name">関数名</param>
            /// <param name = "args">関数の引数</param>
            /// <returns>(Result, Args, Message, Error)</returns>
            public static IFunctionResult Function(string func_name, params object[] args)
            {
                return _AsFunction<Object>(func_name, args);
            }

            /// <summary>
            /// 秀丸マクロの「関数」を実行。関数だけだと返り値が不明な場合にこの<T>付きを使用する。
            /// </summary>
            /// <param name = "func_name">関数名</param>
            /// <param name = "args">関数の引数</param>
            /// <typeparam name="T">String | int | long | IntPtr | double。関数単体だけ確定されない返り値の型を「文字列タイプ」か「整数タイプ」かに振り分け直す。</typeparam>
            /// <returns>(Result, Args, Message, Error)</returns>
            public static IFunctionResult Function<T>(string func_name, params object[] args)
            {
                return _AsFunction<T>(func_name, args);
            }

            public static IFunctionResult _AsFunction<T>(string func_name, params object[] args)
            {
                string funcname = func_name;
                if (funciton_base_random == 0)
                {
                    funciton_base_random = new System.Random().Next(Int16.MaxValue) + 1;

                }

                List<KeyValuePair<string, object>> arg_list = SetMacroVarAndMakeMacroKeyArray(args, funciton_base_random);

                // keyをリスト化
                var arg_keys = new List<String>();
                foreach (var l in arg_list)
                {
                    arg_keys.Add(l.Key);
                }

                // それを「,」で繋げる
                string args_string = String.Join(", ", arg_keys);
                // それを指定の「関数」で実行する形
                string expression = "";

                string result_temp = "";
                Macro.IResult eval_result = new TResult(-1, "", null);
                if (typeof(T) == typeof(int) || typeof(T) == typeof(long) || typeof(T) == typeof(IntPtr) || typeof(T) == typeof(double))
                {
                    expression = $"{funcname}({args_string})";
                    result_temp = "##_tmp_dll_expression_ret";
                    string eval_expresson = result_temp + " = " + expression + ";\n";
                    eval_result = Eval(eval_expresson);
                    expression = result_temp;
                }
                else if (typeof(T) == typeof(String))
                {
                    expression = $"{funcname}({args_string})";
                    result_temp = "$$_tmp_dll_expression_ret";
                    string eval_expresson = result_temp + " = " + expression + ";\n";
                    eval_result = Eval(eval_expresson);
                    expression = result_temp;
                }
                else
                {
                    expression = $"{funcname}({args_string})";
                }
                //----------------------------------------------------------------
                TFunctionResult result = new TFunctionResult(null, "", null, new List<Object>());
                result.Args = new List<object>();

                Object ret = null;
                try
                {
                    ret = Macro.Var[expression]; // この中のGetMethodで例外が発生する可能性あり

                    if (ret.GetType().Name != "String")
                    {
                        if (IntPtr.Size == 4)
                        {
                            result.Result = (Int32)ret + 0; // 確実に複製を
                            result.Message = "";
                            result.Error = null;
                        }
                        else
                        {
                            result.Result = (Int64)ret + 0; // 確実に複製を
                            result.Message = "";
                            result.Error = null;
                        }
                    }
                    else
                    {
                        result.Result = (String)ret + ""; // 確実に複製を
                        result.Message = "";
                        result.Error = null;
                    }

                }
                catch (Exception e)
                {
                    result.Result = null;
                    result.Message = "";
                    result.Error = e;
                }

                if (result_temp.StartsWith("#"))
                {
                    Macro.Var[result_temp] = 0;
                    if (eval_result?.Error != null)
                    {
                        result.Result = null;
                        result.Message = "";
                        result.Error = eval_result.Error;
                    }
                }
                else if (result_temp.StartsWith("$"))
                {
                    Macro.Var[result_temp] = "";
                    if (eval_result?.Error != null)
                    {
                        result.Result = null;
                        result.Message = "";
                        result.Error = eval_result.Error;
                    }
                }

                // 使ったので削除
                for (int ix = 0; ix < arg_list.Count; ix++)
                {
                    var l = arg_list[ix];
                    if (l.Value is Int32 || l.Value is Int64)
                    {
                        result.Args.Add(Macro.Var[l.Key]);
                        Macro.Var[l.Key] = 0;
                    }
                    else if (l.Value is string)
                    {
                        result.Args.Add(Macro.Var[l.Key]);
                        Macro.Var[l.Key] = "";
                    }

                    else if (l.Value.GetType() == new List<int>().GetType() || l.Value.GetType() == new List<long>().GetType() || l.Value.GetType() == new List<IntPtr>().GetType())
                    {
                        result.Args.Add(l.Value);
                        if (l.Value.GetType() == new List<int>().GetType())
                        {
                            List<int> int_list = (List<int>)l.Value;
                            for (int iix = 0; iix < int_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                        else if (l.Value.GetType() == new List<long>().GetType())
                        {
                            List<long> long_list = (List<long>)l.Value;
                            for (int iix = 0; iix < long_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                        else if (l.Value.GetType() == new List<IntPtr>().GetType())
                        {
                            List<IntPtr> ptr_list = (List<IntPtr>)l.Value;
                            for (int iix = 0; iix < ptr_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                    }
                    else if (l.Value.GetType() == new List<String>().GetType())
                    {
                        result.Args.Add(l.Value);
                        List<String> ptr_list = (List<String>)l.Value;
                        for (int iix = 0; iix < ptr_list.Count; iix++)
                        {
                            Macro.Var[l.Key + "[" + iix + "]"] = "";
                        }
                    }
                    else
                    {
                        result.Args.Add(l.Value);
                    }
                }

                return result;
            }

            private static List<KeyValuePair<string, object>> SetMacroVarAndMakeMacroKeyArray(object[] args, int base_random)
            {
                var arg_list = new List<KeyValuePair<String, Object>>();
                int cur_random = new Random().Next(Int16.MaxValue) + 1;
                foreach (var value in args)
                {
                    bool success = false;
                    cur_random++;
                    object normalized_arg = null;
                    // Boolean型であれば、True:1 Flase:0にマッピングする
                    if (value is bool)
                    {
                        success = true;
                        if ((bool)value == true)
                        {
                            normalized_arg = 1;
                        }
                        else
                        {
                            normalized_arg = 0;
                        }
                    }

                    if (value is string || value is StringBuilder)
                    {
                        success = true;
                        normalized_arg = value.ToString();
                    }

                    // 配列の場合を追加
                    if (!success)
                    {
                        if (value.GetType() == new List<int>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<int>((List<int>)value);
                        }
                        if (value.GetType() == new List<long>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<long>((List<long>)value);
                        }
                        if (value.GetType() == new List<IntPtr>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<IntPtr>((List<IntPtr>)value);
                        }
                    }

                    if (!success)
                    {
                        if (value.GetType() == new List<string>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<String>((List<String>)value);
                        }
                    }
                    // 以上配列の場合を追加

                    if (!success)
                    {
                        // 32bit
                        if (IntPtr.Size == 4)
                        {
                            // まずは整数でトライ
                            Int32 itmp = 0;
                            success = Int32.TryParse(value.ToString(), out itmp);

                            if (success == true)
                            {
                                itmp = HmClamp<Int32>(itmp, Int32.MinValue, Int32.MaxValue);
                                normalized_arg = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                Double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int32.MinValue, Int32.MaxValue);
                                    normalized_arg = (Int32)(dtmp);
                                }

                                else
                                {
                                    normalized_arg = 0;
                                }
                            }
                        }

                        // 64bit
                        else
                        {
                            // まずは整数でトライ
                            Int64 itmp = 0;
                            success = Int64.TryParse(value.ToString(), out itmp);

                            if (success == true)
                            {
                                itmp = HmClamp<Int64>(itmp, Int64.MinValue, Int64.MaxValue);
                                normalized_arg = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                Double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int64.MinValue, Int64.MaxValue);
                                    normalized_arg = (Int64)(dtmp);
                                }
                                else
                                {
                                    normalized_arg = 0;
                                }
                            }
                        }
                    }


                    // 成功しなかった
                    if (!success)
                    {
                        normalized_arg = value.ToString();
                    }

                    if (normalized_arg is Int32 || normalized_arg is Int64)
                    {
                        string key = "#AsMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        Macro.Var[key] = normalized_arg;
                    }
                    else if (normalized_arg is string)
                    {
                        string key = "$AsMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        Macro.Var[key] = normalized_arg;
                    }
                    else if (value.GetType() == new List<int>().GetType() || value.GetType() == new List<long>().GetType() || value.GetType() == new List<IntPtr>().GetType())
                    {
                        string key = "$AsIntArrayOfMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        if (value.GetType() == new List<int>().GetType())
                        {
                            List<int> int_list = (List<int>)value;
                            for (int iix = 0; iix < int_list.Count; iix++)
                            {
                                Macro.Var[key + "[" + iix + "]"] = int_list[iix];
                            }
                        }
                        else if (value.GetType() == new List<long>().GetType())
                        {
                            List<long> long_list = (List<long>)value;
                            for (int iix = 0; iix < long_list.Count; iix++)
                            {
                                Macro.Var[key + "[" + iix + "]"] = long_list[iix];
                            }
                        }
                        else if (value.GetType() == new List<IntPtr>().GetType())
                        {
                            List<IntPtr> ptr_list = (List<IntPtr>)value;
                            for (int iix = 0; iix < ptr_list.Count; iix++)
                            {
                                Macro.Var[key + "[" + iix + "]"] = ptr_list[iix];
                            }
                        }
                    }
                    else if (value.GetType() == new List<string>().GetType())
                    {
                        string key = "$AsStrArrayOfMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        List<String> str_list = (List<String>)value;
                        for (int iix = 0; iix < str_list.Count; iix++)
                        {
                            Macro.Var[key + "[" + iix + "]"] = str_list[iix];
                        }
                    }
                }
                return arg_list;
            }


            internal static TMacroVar Var = new TMacroVar();
            internal sealed class TMacroVar
            {
                /// <summary>
                /// 対象の「秀丸マクロ変数名」への読み書き
                /// </summary>
                /// <param name = "var_name">変数のシンボル名</param>
                /// <param name = "value">書き込みの場合、代入する値</param>
                /// <returns>読み取りの場合は、対象の変数の値</returns>
                public Object this[String var_name]
                {
                    get
                    {
                        return GetMethod(var_name);
                    }
                    set
                    {
                        value = SetMethod(var_name, value);
                    }
                }

                private static object SetMethod(string var_name, object value)
                {
                    if (var_name.StartsWith("#"))
                    {
                        Object result = new Object();

                        // Boolean型であれば、True:1 Flase:0にマッピングする
                        if (value is bool)
                        {
                            if ((Boolean)value == true)
                            {
                                value = 1;
                            }
                            else
                            {
                                value = 0;
                            }
                        }

                        // 32bit
                        if (IntPtr.Size == 4)
                        {
                            // まずは整数でトライ
                            Int32 itmp = 0;
                            bool success = Int32.TryParse(value.ToString(), out itmp);

                            if (success == true)
                            {
                                itmp = HmClamp<Int32>(itmp, Int32.MinValue, Int32.MaxValue);
                                result = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                Double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int32.MinValue, Int32.MaxValue);
                                    result = (Int32)(dtmp);
                                }

                                else
                                {
                                    result = 0;
                                }
                            }
                        }

                        // 64bit
                        else
                        {
                            // まずは整数でトライ
                            Int64 itmp = 0;
                            bool success = Int64.TryParse(value.ToString(), out itmp);

                            if (success == true)
                            {
                                itmp = HmClamp<Int64>(itmp, Int64.MinValue, Int64.MaxValue);
                                result = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                Double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int64.MinValue, Int64.MaxValue);
                                    result = (Int64)(dtmp);
                                }
                                else
                                {
                                    result = 0;
                                }
                            }
                        }
                        HmMacroCOMVar.SetVar(var_name, value);
                        HmMacroCOMVar.ClearVar();
                    }

                    else // if (var_name.StartsWith("$")
                    {

                        String result = value.ToString();
                        HmMacroCOMVar.SetVar(var_name, value);
                        HmMacroCOMVar.ClearVar();
                    }

                    return value;
                }

                private static object GetMethod(string var_name)
                {
                    HmMacroCOMVar.ClearVar();
                    Object ret = HmMacroCOMVar.GetVar(var_name);
                    if (ret.GetType().Name != "String")
                    {
                        if (IntPtr.Size == 4)
                        {
                            return (Int32)ret + 0; // 確実に複製を
                        }
                        else
                        {
                            return (Int64)ret + 0; // 確実に複製を
                        }
                    }
                    else
                    {
                        return (String)ret + ""; // 確実に複製を
                    }
                }
            }
        }
    }
}


