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

        // 上一次掃描結果（用來做 diff）
        private Dictionary<int, PortEntry> _last = new();

        public TrayAppContext()
        {
            _tray = new NotifyIcon
            {
                Icon = System.Drawing.SystemIcons.Shield,
                Visible = true,
                Text = "PortWatcher (TCP LISTEN)"
            };

            var menu = new ContextMenuStrip();
            menu.Items.Add("立即掃描", null, (_, __) => ScanAndNotify(forceNotify: true));
            menu.Items.Add("顯示目前列表", null, (_, __) => ShowSnapshot());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("退出", null, (_, __) => ExitApp());

            _tray.ContextMenuStrip = menu;
            _tray.DoubleClick += (_, __) => ShowSnapshot();

            // 先掃一次
            ScanAndNotify(forceNotify: false);

            _timer = new System.Windows.Forms.Timer { Interval = 5000 };
            _timer.Tick += (_, __) => ScanAndNotify(forceNotify: false);
            _timer.Start();
        }

        private void ExitApp()
        {
            _timer.Stop();
            _tray.Visible = false;
            _tray.Dispose();
            Application.Exit();
        }

        private void ShowSnapshot()
        {
            var now = SafeGetListeningPorts();
            var sb = new StringBuilder();

            sb.AppendLine($"TCP LISTEN Ports: {now.Count}");
            sb.AppendLine(new string('-', 48));

            foreach (var kv in now.OrderBy(k => k.Key))
            {
                var p = kv.Value;
                sb.AppendLine($"{p.Port,5}  PID {p.Pid,6}  {p.ProcessName}");
            }

            MessageBox.Show(sb.ToString(), "PortWatcher - Snapshot", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            // 更新 last
            _last = now;

            if (!forceNotify && added.Count == 0 && removed.Count == 0)
                return;

            var msg = new StringBuilder();

            if (added.Count > 0)
            {
                msg.AppendLine("新增 LISTEN：");
                foreach (var port in added.Take(10))
                {
                    var e = now[port];
                    msg.AppendLine($"  {e.Port}  {e.ProcessName} (PID {e.Pid})");
                }
                if (added.Count > 10) msg.AppendLine($"  ... 另有 {added.Count - 10} 個");
            }

            if (removed.Count > 0)
            {
                if (msg.Length > 0) msg.AppendLine();
                msg.AppendLine("釋放 LISTEN：");
                foreach (var port in removed.Take(10))
                {
                    msg.AppendLine($"  {port}");
                }
                if (removed.Count > 10) msg.AppendLine($"  ... 另有 {removed.Count - 10} 個");
            }

            if (forceNotify && added.Count == 0 && removed.Count == 0)
                msg.AppendLine("沒有變化（仍在監控中）。");

            _tray.BalloonTipTitle = "PortWatcher";
            _tray.BalloonTipText = msg.ToString().Trim();
            _tray.ShowBalloonTip(3000);
        }

        private static Dictionary<int, PortEntry> SafeGetListeningPorts()
        {
            try
            {
                var list = TcpTable.GetTcpListenersWithPid();
                var dict = new Dictionary<int, PortEntry>();

                foreach (var row in list)
                {
                    // 只取 LISTEN
                    if (row.State != MibTcpState.LISTEN) continue;

                    var port = row.LocalPort;
                    var pid = row.Pid;

                    // 同一個 port 可能多個（很少見），這裡以第一個為主
                    if (dict.ContainsKey(port)) continue;

                    var pname = "Unknown";
                    try
                    {
                        using var p = Process.GetProcessById(pid);
                        pname = p.ProcessName;
                    }
                    catch
                    {
                        // 有些系統/權限情況會查不到
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
    }

    internal record PortEntry(int Port, int Pid, string ProcessName);

    // ====== 下面是「取得 TCP table + PID」的底層 API（Windows 原生）======

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
        public uint localPort;  // network order in high 2 bytes
        public uint remoteAddr;
        public uint remotePort; // network order in high 2 bytes
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

        public static List<MIB_TCPROW_OWNER_PID> GetTcpListenersWithPid()
        {
            int buffSize = 0;
            // 第一次呼叫拿 buffer size
            uint ret = GetExtendedTcpTable(IntPtr.Zero, ref buffSize, true, AF_INET, TCP_TABLE_CLASS.TCP_TABLE_OWNER_PID_ALL, 0);
            IntPtr buff = Marshal.AllocHGlobal(buffSize);

            try
            {
                ret = GetExtendedTcpTable(buff, ref buffSize, true, AF_INET, TCP_TABLE_CLASS.TCP_TABLE_OWNER_PID_ALL, 0);
                if (ret != 0) throw new Exception("GetExtendedTcpTable failed: " + ret);

                // 前 4 bytes 是 entry count
                int offset = 0;
                int numEntries = Marshal.ReadInt32(buff);
                offset += 4;

                var rowSize = Marshal.SizeOf<MIB_TCPROW_OWNER_PID>();
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
