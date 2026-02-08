using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace PortWatcher
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new TrayAppContext());
        }
    }

    internal sealed class TrayAppContext : ApplicationContext
    {
        private readonly NotifyIcon _tray;
        private readonly System.Windows.Forms.Timer _timer;
        private ContextMenuStrip _menu = null!;

        private Dictionary<int, PortEntry> _last = new();

        private AppLang _lang = AppLang.Zh;
        private SortMode _sort = SortMode.Port;

        private enum AppLang { Zh, En }
        private enum SortMode { Port, Pid, Process }

        public TrayAppContext()
        {
            _tray = new NotifyIcon
            {
                Icon = LoadTrayIcon()
            };
            _tray.ShowBalloonTip(
                3000,
                "PortWatcher",
                "PortWatcher 已在系統匣執行（右下角），右鍵可退出",
                ToolTipIcon.Info
            );

            BuildMenu();
            UpdateTrayText();

            _tray.DoubleClick += (_, __) => ShowSnapshot();

            ScanAndNotify(forceNotify: false);

            _timer = new System.Windows.Forms.Timer { Interval = 5000 };
            _timer.Tick += (_, __) => ScanAndNotify(forceNotify: false);
            _timer.Start();
        }

        private static System.Drawing.Icon LoadTrayIcon()
        {
            var asm = typeof(TrayAppContext).Assembly;

            const string resName = "PortWatcher.assets.portwatcher.ico";

            using var stream = asm.GetManifestResourceStream(resName);
            if (stream is null)
                return System.Drawing.SystemIcons.Shield; // 找不到就用預設，避免秒退

            return new System.Drawing.Icon(stream);
        }


        private void BuildMenu()
        {
            _menu?.Dispose();
            _menu = new ContextMenuStrip();

            _menu.Items.Add(T("Menu.ScanNow"), null, (_, __) => ScanAndNotify(forceNotify: true));
            _menu.Items.Add(T("Menu.Snapshot"), null, (_, __) => ShowSnapshot());
            _menu.Items.Add(new ToolStripSeparator());

            var langMenu = new ToolStripMenuItem(T("Menu.Language"));
            var miZh = new ToolStripMenuItem(T("Menu.LangZh"), null, (_, __) => { _lang = AppLang.Zh; RebuildUI(); });
            var miEn = new ToolStripMenuItem(T("Menu.LangEn"), null, (_, __) => { _lang = AppLang.En; RebuildUI(); });
            langMenu.DropDownItems.Add(miZh);
            langMenu.DropDownItems.Add(miEn);
            _menu.Items.Add(langMenu);

            var sortMenu = new ToolStripMenuItem(T("Menu.Sort"));
            var miSortPort = new ToolStripMenuItem(T("Menu.SortPort"), null, (_, __) => { _sort = SortMode.Port; UpdateChecks(); });
            var miSortPid = new ToolStripMenuItem(T("Menu.SortPid"), null, (_, __) => { _sort = SortMode.Pid; UpdateChecks(); });
            var miSortProc = new ToolStripMenuItem(T("Menu.SortProcess"), null, (_, __) => { _sort = SortMode.Process; UpdateChecks(); });
            sortMenu.DropDownItems.Add(miSortPort);
            sortMenu.DropDownItems.Add(miSortPid);
            sortMenu.DropDownItems.Add(miSortProc);
            _menu.Items.Add(sortMenu);

            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(T("Menu.Exit"), null, (_, __) => ExitApp());

            _tray.ContextMenuStrip = _menu;

            UpdateChecks();
        }

        private void RebuildUI()
        {
            BuildMenu();
            UpdateTrayText();
        }

        private void UpdateTrayText()
        {
            _tray.Text = _lang == AppLang.Zh ? "PortWatcher（TCP LISTEN）" : "PortWatcher (TCP LISTEN)";
        }

        private void UpdateChecks()
        {
            foreach (ToolStripItem item in _menu.Items)
            {
                if (item is not ToolStripMenuItem mi) continue;

                if (mi.Text == T("Menu.Language"))
                {
                    foreach (ToolStripItem subItem in mi.DropDownItems)
                    {
                        if (subItem is not ToolStripMenuItem sub) continue;
                        sub.Checked =
                            (_lang == AppLang.Zh && sub.Text == T("Menu.LangZh")) ||
                            (_lang == AppLang.En && sub.Text == T("Menu.LangEn"));
                    }
                }

                if (mi.Text == T("Menu.Sort"))
                {
                    foreach (ToolStripItem subItem in mi.DropDownItems)
                    {
                        if (subItem is not ToolStripMenuItem sub) continue;
                        sub.Checked =
                            (_sort == SortMode.Port && sub.Text == T("Menu.SortPort")) ||
                            (_sort == SortMode.Pid && sub.Text == T("Menu.SortPid")) ||
                            (_sort == SortMode.Process && sub.Text == T("Menu.SortProcess"));
                    }
                }
            }
        }

        private void ExitApp()
        {
            _timer.Stop();
            _tray.Visible = false;
            _tray.Dispose();
            _menu.Dispose();
            Application.Exit();
        }

        private void ShowSnapshot()
        {
            var now = SafeGetListeningPorts();

            var ordered = _sort switch
            {
                SortMode.Port => now.Values.OrderBy(x => x.Port),
                SortMode.Pid => now.Values.OrderBy(x => x.Pid).ThenBy(x => x.Port),
                SortMode.Process => now.Values
                    .OrderBy(x => x.ProcessName, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(x => x.Port),
                _ => now.Values.OrderBy(x => x.Port)
            };

            // 
            int portW = Math.Max(4, ordered.Any() ? ordered.Max(x => x.Port.ToString().Length) : 4);
            int pidW = Math.Max(3, ordered.Any() ? ordered.Max(x => x.Pid.ToString().Length) : 3);
            int nameW = Math.Min(40, Math.Max(8, ordered.Any() ? ordered.Max(x => (x.ProcessName ?? "").Length) : 8)); // 最多 40

            string Trunc(string s, int w)
            {
                s ??= "";
                if (s.Length <= w) return s;
                return s.Substring(0, Math.Max(0, w - 1)) + "…";
            }

            var sb = new StringBuilder();
            sb.AppendLine($"{T("Snapshot.Header")}: {now.Count}");
            sb.AppendLine(new string('-', portW + pidW + nameW + 10));
            sb.AppendLine($"{Pad("PORT", portW)}  {Pad("PID", pidW)}  {Pad("PROCESS", nameW)}");
            sb.AppendLine(new string('-', portW + pidW + nameW + 10));

            foreach (var p in ordered)
            {
                sb.AppendLine($"{p.Port.ToString().PadLeft(portW)}  {p.Pid.ToString().PadLeft(pidW)}  {Trunc(p.ProcessName, nameW).PadRight(nameW)}");
            }

            var form = new Form
            {
                Text = $"{T("Title")} - {T("Menu.Snapshot")}",
                StartPosition = FormStartPosition.CenterScreen,
                Width = 720,
                Height = 520
            };

            var box = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Consolas", 10f),
                Text = sb.ToString()
            };

            // Ctrl+A 
            box.KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.A)
                {
                    box.SelectAll();
                    e.SuppressKeyPress = true;
                }
            };

            form.Controls.Add(box);
            form.Show(); // 
        }

        private static string Pad(string s, int width)
        {
            if (s.Length >= width) return s;
            return s.PadRight(width);
        }


        private void ScanAndNotify(bool forceNotify)
        {
            var now = SafeGetListeningPorts();

            if (_last.Count == 0 && !forceNotify)
            {
                _last = now;
                return;
            }

            var added = now.Keys.Except(_last.Keys).OrderBy(x => x).ToList();
            var removed = _last.Keys.Except(now.Keys).OrderBy(x => x).ToList();

            _last = now;

            if (!forceNotify && added.Count == 0 && removed.Count == 0)
                return;

            var msg = new StringBuilder();

            if (added.Count > 0)
            {
                msg.AppendLine(T("Notify.Added"));
                foreach (var port in added.Take(10))
                {
                    var e = now[port];
                    msg.AppendLine($"  {e.Port}  {e.ProcessName} (PID {e.Pid})");
                }
                if (added.Count > 10) msg.AppendLine(T("Notify.More", added.Count - 10));
            }

            if (removed.Count > 0)
            {
                if (msg.Length > 0) msg.AppendLine();
                msg.AppendLine(T("Notify.Removed"));
                foreach (var port in removed.Take(10))
                {
                    msg.AppendLine($"  {port}");
                }
                if (removed.Count > 10) msg.AppendLine(T("Notify.More", removed.Count - 10));
            }

            if (forceNotify && added.Count == 0 && removed.Count == 0)
            {
                msg.AppendLine(T("Notify.NoChange"));
            }

            _tray.BalloonTipTitle = T("Title");
            _tray.BalloonTipText = msg.ToString().Trim();
            _tray.ShowBalloonTip(3000);
        }

        private static Dictionary<int, PortEntry> SafeGetListeningPorts()
        {
            try
            {
                var rows = TcpTable.GetTcpTableOwnerPidAll();
                var dict = new Dictionary<int, PortEntry>();

                foreach (var row in rows)
                {
                    if (row.State != MibTcpState.LISTEN) continue;

                    var port = row.LocalPort;
                    var pid = row.Pid;

                    if (port <= 0 || port > 65535) continue;
                    if (dict.ContainsKey(port)) continue;

                    var pname = "Unknown";
                    try
                    {
                        using var p = Process.GetProcessById(pid);
                        pname = p.ProcessName;
                    }
                    catch
                    {
                        // 權限/瞬間結束的程序可能查不到
                    }

                    dict[port] = new PortEntry(port, pid, pname);
                }

                return dict;
            }
            catch
            {
                return new Dictionary<int, PortEntry>();
            }
        }

        private string T(string key, int n = 0)
        {
            return (_lang, key) switch
            {
                (AppLang.Zh, "Title") => "PortWatcher",
                (AppLang.En, "Title") => "PortWatcher",

                (AppLang.Zh, "Menu.ScanNow") => "立即掃描",
                (AppLang.En, "Menu.ScanNow") => "Scan now",

                (AppLang.Zh, "Menu.Snapshot") => "顯示目前列表",
                (AppLang.En, "Menu.Snapshot") => "Show snapshot",

                (AppLang.Zh, "Menu.Exit") => "退出",
                (AppLang.En, "Menu.Exit") => "Exit",

                (AppLang.Zh, "Menu.Language") => "語言",
                (AppLang.En, "Menu.Language") => "Language",

                (AppLang.Zh, "Menu.LangZh") => "中文",
                (AppLang.En, "Menu.LangZh") => "Chinese",

                (AppLang.Zh, "Menu.LangEn") => "English",
                (AppLang.En, "Menu.LangEn") => "English",

                (AppLang.Zh, "Menu.Sort") => "排序",
                (AppLang.En, "Menu.Sort") => "Sort",

                (AppLang.Zh, "Menu.SortPort") => "依 Port",
                (AppLang.En, "Menu.SortPort") => "By Port",

                (AppLang.Zh, "Menu.SortPid") => "依 PID",
                (AppLang.En, "Menu.SortPid") => "By PID",

                (AppLang.Zh, "Menu.SortProcess") => "依 程式名稱",
                (AppLang.En, "Menu.SortProcess") => "By Process name",

                (AppLang.Zh, "Snapshot.Header") => "TCP LISTEN Ports",
                (AppLang.En, "Snapshot.Header") => "TCP LISTEN Ports",

                (AppLang.Zh, "Notify.Added") => "新增 LISTEN：",
                (AppLang.En, "Notify.Added") => "Listening added:",

                (AppLang.Zh, "Notify.Removed") => "釋放 LISTEN：",
                (AppLang.En, "Notify.Removed") => "Listening removed:",

                (AppLang.Zh, "Notify.NoChange") => "沒有變化（仍在監控中）。",
                (AppLang.En, "Notify.NoChange") => "No change (still monitoring).",

                (AppLang.Zh, "Notify.More") => $"  ... 另有 {n} 個",
                (AppLang.En, "Notify.More") => $"  ... plus {n} more",

                _ => key
            };
        }
    }

    internal record PortEntry(int Port, int Pid, string ProcessName);

    internal enum MibTcpState : uint
    {
        CLOSED = 1,
        LISTEN = 2,
        SYN_SENT = 3,
        SYN_RECEIVED = 4,
        ESTABLISHED = 5,
        FIN_WAIT1 = 6,
        FIN_WAIT2 = 7,
        CLOSE_WAIT = 8,
        CLOSING = 9,
        LAST_ACK = 10,
        TIME_WAIT = 11,
        DELETE_TCB = 12
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MIB_TCPROW_OWNER_PID
    {
        public MibTcpState state;
        public uint localAddr;
        public uint localPort;
        public uint remoteAddr;
        public uint remotePort;
        public uint owningPid;

        public int LocalPort => (ushort)IPAddress.NetworkToHostOrder((short)localPort);
        public int Pid => (int)owningPid;
        public MibTcpState State => state;
    }

    internal static class TcpTable
    {
        private const int AF_INET = 2;

        private enum TCP_TABLE_CLASS
        {
            TCP_TABLE_BASIC_LISTENER,
            TCP_TABLE_BASIC_CONNECTIONS,
            TCP_TABLE_BASIC_ALL,
            TCP_TABLE_OWNER_PID_LISTENER,
            TCP_TABLE_OWNER_PID_CONNECTIONS,
            TCP_TABLE_OWNER_PID_ALL,
            TCP_TABLE_OWNER_MODULE_LISTENER,
            TCP_TABLE_OWNER_MODULE_CONNECTIONS,
            TCP_TABLE_OWNER_MODULE_ALL
        }

        [DllImport("iphlpapi.dll", SetLastError = true)]
        private static extern uint GetExtendedTcpTable(
            IntPtr pTcpTable,
            ref int dwOutBufLen,
            bool sort,
            int ipVersion,
            TCP_TABLE_CLASS tblClass,
            uint reserved);

        public static List<MIB_TCPROW_OWNER_PID> GetTcpTableOwnerPidAll()
        {
            int buffSize = 0;
            uint ret = GetExtendedTcpTable(IntPtr.Zero, ref buffSize, true, AF_INET, TCP_TABLE_CLASS.TCP_TABLE_OWNER_PID_ALL, 0);

            IntPtr buff = Marshal.AllocHGlobal(buffSize);
            try
            {
                ret = GetExtendedTcpTable(buff, ref buffSize, true, AF_INET, TCP_TABLE_CLASS.TCP_TABLE_OWNER_PID_ALL, 0);
                if (ret != 0) throw new Exception("GetExtendedTcpTable failed: " + ret);

                int offset = 0;
                int numEntries = Marshal.ReadInt32(buff);
                offset += 4;

                int rowSize = Marshal.SizeOf<MIB_TCPROW_OWNER_PID>();
                var list = new List<MIB_TCPROW_OWNER_PID>(numEntries);

                for (int i = 0; i < numEntries; i++)
                {
                    IntPtr rowPtr = IntPtr.Add(buff, offset);
                    var row = Marshal.PtrToStructure<MIB_TCPROW_OWNER_PID>(rowPtr);
                    list.Add(row);
                    offset += rowSize;
                }

                return list;
            }
            finally
            {
                Marshal.FreeHGlobal(buff);
            }
        }
    }
}
