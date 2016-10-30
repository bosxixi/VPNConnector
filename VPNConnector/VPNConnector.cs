using DotRas;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace VPNConnector
{
    public class VPNConnector
    {
        public VPNConnector(string serverAddress, string userName, string passWord)
           : this(serverAddress, serverAddress, userName, passWord, Protocol.SSTP)
        {

        }
        public VPNConnector(string serverAddress, string connectionName, string userName, string passWord)
            : this(serverAddress, connectionName, userName, passWord, Protocol.SSTP)
        {

        }

        public VPNConnector(string serverAddress, string connectionName, string userName, string passWord, Protocol protocol)
        {
            this.serverAddress = serverAddress;
            this.connectionName = connectionName;
            this.userName = userName;
            this.passWord = passWord;
            this.protocol = protocol;
            this.rasDialFileName = Path.Combine(WinDir, "rasdial.exe"); 
        }

        private static string WinDir = Environment.GetFolderPath(Environment.SpecialFolder.System);
        private string rasDialFileName;

        public string RasDialFileName
        {
            get { return rasDialFileName; }
            set {
                if (File.Exists(value))
                {
                    rasDialFileName = value;
                }

                throw new FileNotFoundException();
            }
        }

        private readonly string serverAddress;
        private readonly string connectionName;
        private readonly string userName;
        private readonly string passWord;
        private readonly Protocol protocol;
        private readonly static string allUserPhoneBookPath = RasPhoneBook.GetPhoneBookPath(RasPhoneBookType.AllUsers);
        public bool IsActive
        {
            get
            {
                using (Process myProcess = new Process())
                {
                    myProcess.StartInfo.CreateNoWindow = true;
                    myProcess.StartInfo.UseShellExecute = false;
                    myProcess.StartInfo.RedirectStandardInput = true;
                    myProcess.StartInfo.RedirectStandardOutput = true;
                    myProcess.StartInfo.FileName = "cmd.exe";
                    myProcess.Start();
                    myProcess.StandardInput.WriteLine("ipconfig");
                    myProcess.StandardInput.WriteLine("exit");
                    myProcess.WaitForExit();

                    string content = myProcess.StandardOutput.ReadToEnd();
                    if (content.Contains("0.0.0.0"))
                    {
                        return true;
                    }

                    return false;
                }

            }
        }

        public RasDevice RasDevice
        {
            get
            {
                var name = Enum.GetName(typeof(Protocol), this.protocol);
                var rasDevice = RasDevice.GetDevices().FirstOrDefault(c => c.Name.Contains(name));
                if (rasDevice == null)
                {
                    throw new Exception("No device found.");
                }
                return rasDevice;
            }
        }

        public RasVpnStrategy RasVpnStrategy
        {
            get
            {
                if (protocol == Protocol.SSTP)
                {
                    return RasVpnStrategy.SstpFirst;
                }
                else
                {
                    return RasVpnStrategy.IkeV2First;
                }
            }
        }

        public bool WaitUntilActive(int timeOut = 10)
        {
            for (int i = 0; i < timeOut; i++)
            {
                if (!this.IsActive)
                {
                    Thread.Sleep(1000);
                }
                else
                {
                    return true;
                }
            }

            return false;
        }

        public bool WaitUntilInActive(int timeOut = 10)
        {
            for (int i = 0; i < timeOut; i++)
            {
                if (this.IsActive)
                {
                    Thread.Sleep(1000);
                }
                else
                {
                    return true;
                }
            }

            return false;
        }

        public bool TryConnect()
        {
            try
            {
                string args = $"{connectionName} {userName} {passWord}";
                ProcessStartInfo myProcess = new ProcessStartInfo(rasDialFileName, args);
                myProcess.CreateNoWindow = true;
                myProcess.UseShellExecute = false;
                Process.Start(myProcess);
            }
            catch (Exception Ex)
            {
                Debug.Assert(false, Ex.ToString());
            }

            WaitUntilActive();
            if (IsActive)
            {
                return true;
            }

            return false;
        }

        public bool TryDisconnect()
        {
            try
            {
                string args = $@"{connectionName} /d";
                ProcessStartInfo myProcess = new ProcessStartInfo(rasDialFileName, args);
                myProcess.CreateNoWindow = true;
                myProcess.UseShellExecute = false;
                Process.Start(myProcess);
            }
            catch (Exception Ex)
            {
                Debug.Assert(false, Ex.ToString());
            }

            WaitUntilInActive();
            if (!IsActive)
            {
                return true;
            }

            return false;
        }

        public void CreateOrUpdate()
        {
            using (var dialer = new RasDialer())
            using (var allUsersPhoneBook = new RasPhoneBook())
            {
                allUsersPhoneBook.Open(true);
                if (allUsersPhoneBook.Entries.Contains(connectionName))
                {
                    allUsersPhoneBook.Entries[connectionName].PhoneNumber = connectionName;
                    allUsersPhoneBook.Entries[connectionName].VpnStrategy = RasVpnStrategy;
                    allUsersPhoneBook.Entries[connectionName].Device = RasDevice;
                    allUsersPhoneBook.Entries[connectionName].Update();
                }
                else
                {
                    RasEntry entry = RasEntry.CreateVpnEntry(connectionName, serverAddress, RasVpnStrategy, RasDevice);
                    allUsersPhoneBook.Entries.Add(entry);
                    dialer.EntryName = connectionName;
                    dialer.PhoneBookPath = allUserPhoneBookPath;
                }
            }
        }

        public void TryDelete()
        {
            using (var dialer = new RasDialer())
            using (var allUsersPhoneBook = new RasPhoneBook())
            {
                allUsersPhoneBook.Open(true);
                if (allUsersPhoneBook.Entries.Contains(connectionName))
                {
                    TryDisconnect();
                    WaitUntilInActive();
                    allUsersPhoneBook.Entries.Remove(connectionName);
                }
            }
        }
    }
}
