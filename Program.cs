using Gameloop.Vdf;
using Gameloop.Vdf.Linq;
using Microsoft.VisualBasic.FileIO;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 얼불춤_모드_자동설치기
{
    class Program
    {
        private const string STEAM_APP_ID = "977950";
        private const string STEAM_32 = @"Software\Valve\Steam";
        private const string STEAM_64 = @"Software\Wow6432Node\Valve\Steam";
        private const string DOWNLOAD_URL = "https://www.dropbox.com/s/wz8x8e4onjdfdbm/UnityModManager.zip?dl=1";
        private const string SPREADSHEET_URL_START = "https://docs.google.com/spreadsheets/d/";
        private const string SPREADSHEET_URL_END = "/gviz/tq?tqx=out:json&tq&gid=";
        private const string KEY = "1QcrRL6LAs8WxJj_hFsEJa3CLM5g3e8Ya0KQlRKXwdlU";
        private const string GID = "664160099";
        private const string SPREADSHEET_URL = SPREADSHEET_URL_START + KEY + SPREADSHEET_URL_END + GID;
        private const string MOD_URL_START = "https://bot.adofai.gg/api/mods/";
        private const string MOD_URL_END = "?download=true";

        private static readonly List<Dictionary<string, string>> localizations = new List<Dictionary<string, string>>
        {
            new Dictionary<string, string>
            {
                { "wrong_input", "다시 입력하세요." },
                { "network_not_connected", "네트워크가 연결되지 않았습니다." },
                { "steam_not_found", "스팀 폴더를 찾을 수 없습니다! 얼불춤이 설치된 위치를 직접 입력해주세요." },
                { "file_not_found", "필요한 파일이 없습니다." },
                { "adofai_not_found", "얼불춤 폴더를 찾을 수 없습니다! 얼불춤이 설치된 위치를 직접 입력해주세요." },
                { "adofai_running", "얼불춤이 켜져있습니다! 얼불춤을 끈 후 다시 실행해보세요." },

                { "umm_already_installed", "모드 매니저가 이미 설치되어 있습니다." },
                { "downloading_umm", "모드 매니저 다운로드중..." },
                { "installing_umm", "모드매니저를 설치합니다. 엔터 키를 누르고 기다리세요." },
                { "installed_umm", "모드 매니저 설치 완료!" },

                { "adofaigg", "모드들을 직접 설치하려면 아래 사이트를 참고하세요.\nhttps://www.notion.so/tjwogud/2e12640062ce45228916b1740328d863" },

                { "download_recommended", "추천 모드를 설치하시겠습니까? 설치하려면 y를 입력하세요." },
                { "loading_recommended", "추천 모드 목록 로드중..." },
                { "select_mod", "위 목록 중 설치할 모드 번호를 입력하세요." },
                { "downloading_mod", "모드 다운로드중..." },
                { "installing_mod", "모드 설치중..." },
                { "installed_mod", "모드 설치 완료!" },
                { "wrong_mod", "모드를 설치하지 못했습니다. 개발자에게 제보해주세요." }
            },
            new Dictionary<string, string>
            {
                { "wrong_input", "Please enter again." },
                { "network_not_connected", "Network is not connected." },
                { "steam_not_found", "Can't detect Steam! Please enter adofai path." },
                { "file_not_found", "Needed files are missing." },
                { "adofai_not_found", "Can't detect adofai! Please enter its path." },
                { "adofai_running", "Adofai is running! Please turn off it and try again." },

                { "umm_already_installed", "Mod Manager is already installed." },
                { "downloading_umm", "Downloading Mod Manager..." },
                { "installing_umm", "Installing Mod Manager. Press enter and wait." },
                { "installed_umm", "Installed Mod Manager!" },

                { "adofaigg", "Check below site to install mods manually.\nhttps://tjwogud.notion.site/015e915f69c54d71a338ca1190136a8c" },

                { "download_recommended", "Would you install recommended mods? Enter y to install them." },
                { "loading_recommended", "Loading recommanded mods..." },
                { "select_mod", "Enter the number of mod you want to install." },
                { "downloading_mod", "Downloading mod..." },
                { "installing_mod", "Installing mod..." },
                { "installed_mod", "Installed mod!" },
                { "wrong_mod", "Can't install mod. Please report it to dev." }
            }
        };

        static void Main(string[] args)
        {
            Console.Clear();

            Console.WriteLine("1. 한글\n2. English");
            int lang = 0;
            while (true)
            {
                Console.Write("> ");
                if (int.TryParse(Console.ReadLine(), out int v) && (v == 1 || v == 2))
                {
                    lang = v;
                    break;
                }
                else
                    Console.WriteLine(localizations[lang]["wrong_input"]);
            }
            lang--;

            if (!HasNetworkConnection())
            {
                Console.WriteLine(localizations[lang]["network_not_connected"]);
                return;
            }

            RegistryKey key = Registry.LocalMachine.OpenSubKey(Environment.Is64BitOperatingSystem ? STEAM_64 : STEAM_32);
            object value = key?.GetValue("InstallPath");
            key.Dispose();
            if (value == null)
            {
                Console.WriteLine(localizations[lang]["steam_not_found"]);
                Console.Write("> ");
                value = Console.ReadLine();
            }

            string vdfPath = Path.Combine(value as string, "steamapps", "libraryfolders.vdf");
            if (!File.Exists(vdfPath))
            {
                Console.WriteLine(localizations[lang]["file_not_found"]);
                return;
            }
            string libraryfolders = File.ReadAllText(vdfPath);
            string path = null;
            foreach (VToken token1 in VdfConvert.Deserialize(libraryfolders).Value)
            {
                VToken token2 = (token1 as VProperty).Value;
                if (token2["apps"].Cast<VProperty>().Any(t => t.Key == STEAM_APP_ID))
                {
                    path = token2["path"].Value<string>();
                }
            }
            if (path == null)
            {
                Console.WriteLine(localizations[lang]["adofai_not_found"]);
                Console.Write("> ");
                path = Console.ReadLine();
            }
            else
                path = Path.Combine(path, "steamapps", "common", "A Dance of Fire and Ice");
            string modDir = Path.Combine(path, "Mods");
            if (!Directory.Exists(modDir))
                Directory.CreateDirectory(modDir);
            Console.WriteLine();

            if (!File.Exists(Path.Combine(path, "A Dance of Fire and Ice_Data", "Managed", "UnityModManager", "UnityModManager.dll")))
            {
                Console.WriteLine(localizations[lang]["downloading_umm"]);
                int prevProg = -1;
                int prog = 0;
                bool end = false;
                Task.Run(() =>
                {
                    WebClient client = new WebClient();
                    client.Encoding = Encoding.UTF8;
                    client.DownloadProgressChanged += (sender, e) => prog = e.ProgressPercentage;
                    client.DownloadFileCompleted += (sender, e) => end = true;
                    client.DownloadFileCompleted += (sender, e) => client.Dispose();
                    client.DownloadFileAsync(new Uri(DOWNLOAD_URL), "_umm.zip");
                });
                while (!end)
                {
                    if (prevProg != prog)
                    {
                        Console.Write("\r[");
                        for (int i = 0; i < 20; i++)
                        {
                            if (prog >= i * 5)
                                Console.Write("/");
                            else
                                Console.Write(" ");
                        }
                        Console.Write($"] {prog}%");
                    }
                    prevProg = prog;
                    Thread.Sleep(10);
                }
                Console.WriteLine("\r[////////////////////] 100%");
                Console.WriteLine();

                if (Directory.Exists("_UnityModManager"))
                    Directory.Delete("_UnityModManager", true);
                if (Directory.Exists("UnityModManager"))
                    Directory.Delete("UnityModManager", true);
                ZipFile.ExtractToDirectory("_umm.zip", "_UnityModManager");
                Directory.Move(Path.Combine("_UnityModManager", "UnityModManagerInstaller"), "UnityModManager");
                Directory.Delete("_UnityModManager");
                File.Delete("_umm.zip");

                string ummPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "UnityModManagerNet");
                string paramsPath = Path.Combine(ummPath, "Params.xml");
                if (!Directory.Exists(ummPath))
                {
                    Directory.CreateDirectory(ummPath);
                }
                else
                    File.Delete(paramsPath);
                try
                {
                    File.WriteAllText(paramsPath,
    $@"<?xml version=""1.0"" encoding=""utf-8""?>
<Param xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <LastSelectedGame>A Dance of Fire and Ice</LastSelectedGame>
  <GameParams>
    <GameParam Name=""A Dance of Fire and Ice"">
      <Path>{path}</Path>
      <InstallType>Assembly</InstallType>
    </GameParam>
  </GameParams>
</Param>");
                }
                catch (Exception)
                {
                }
                Console.WriteLine(localizations[lang]["installing_umm"]);
                Console.ReadLine();
                Console.Clear();
                Process umm = new Process();
                umm.StartInfo.FileName = Path.Combine("UnityModManager", "Console.exe");
                umm.StartInfo.UseShellExecute = false;
                umm.StartInfo.RedirectStandardOutput = true;
                umm.StartInfo.RedirectStandardError = true;
                umm.Start();

                SendKeys.SendWait("{Enter}");
                bool install = true;
                while (true)
                {
                    string line = umm.StandardOutput.ReadLine();
                    if (line == null)
                        continue;
                    if (line.Trim().EndsWith("I. Install"))
                        install = true;
                    else if (line.Trim().EndsWith("D. Delete"))
                        install = false;
                    else if (line.Trim().EndsWith("R. Restore"))
                    {
                        if (install)
                            SendKeys.SendWait("i{Enter}{Enter}");
                        else
                        {
                            SendKeys.SendWait("{Enter}{Enter}");
                            break;
                        }
                    }
                    //Console.Clear();
                }
                Console.Clear();
                Console.WriteLine(localizations[lang]["installed_umm"]);
                Console.WriteLine();
            }
            else
            {
                Console.WriteLine(localizations[lang]["umm_already_installed"]);
                Console.WriteLine();
            }

            Console.WriteLine(localizations[lang]["adofaigg"]);
            Console.WriteLine();
            Console.WriteLine(localizations[lang]["download_recommended"]);
            Console.Write("> ");
            if (Console.ReadLine()?.ToLower() == "y")
            {
                Console.WriteLine();
                Console.WriteLine(localizations[lang]["loading_recommended"]);

                WebClient client = new WebClient();
                client.Encoding = Encoding.UTF8;
                string data = client.DownloadString(SPREADSHEET_URL);
                client.Dispose();
                data = data.Substring(data.IndexOf("(") + 1, data.IndexOf(");") - (data.IndexOf("(") + 1));
                JArray array = (JArray)JsonConvert.DeserializeObject<JObject>(data)["table"]["rows"];
                Console.Clear();
                while (true)
                {
                    for (int i = 0; i < array.Count; i++)
                    {
                        JArray mod = (JArray)array[i]["c"];
                        Console.WriteLine($"{i + 1}. {mod[0]["v"]}\n   {mod[lang + 1]["v"]}\n");
                    }
                    Console.WriteLine();
                    Console.WriteLine(localizations[lang]["select_mod"]);
                    Console.Write("> ");
                    string input = Console.ReadLine();
                    Console.Clear();
                    if (int.TryParse(input, out int index) && index > 0 && index <= array.Count)
                    {
                        string name = (string)array[index - 1]["c"][0]["v"];
                        string url = MOD_URL_START + name + MOD_URL_END;
                        Console.WriteLine(localizations[lang]["downloading_mod"]);

                        int prevProg = -1;
                        int prog = 0;
                        bool end = false;
                        Task.Run(() =>
                        {
                            client = new WebClient();
                            client.Encoding = Encoding.UTF8;
                            client.DownloadProgressChanged += (sender, e) => prog = e.ProgressPercentage;
                            client.DownloadFileCompleted += (sender, e) => end = true;
                            client.DownloadFileCompleted += (sender, e) => client.Dispose();
                            client.DownloadFileAsync(new Uri(url), $"_{name}.zip");
                        });
                        while (!end)
                        {
                            if (prevProg != prog)
                            {
                                Console.Write("\r[");
                                for (int i = 0; i < 20; i++)
                                {
                                    if (prog >= i * 5)
                                        Console.Write("/");
                                    else
                                        Console.Write(" ");
                                }
                                Console.Write($"] {prog}%");
                            }
                            prevProg = prog;
                            Thread.Sleep(10);
                        }
                        Console.WriteLine("\r[////////////////////] 100%");
                        Console.WriteLine();
                        Console.WriteLine(localizations[lang]["installing_mod"]);
                        ZipFile.ExtractToDirectory($"_{name}.zip", name);
                        File.Delete($"_{name}.zip");
                        string modPath = name;
                        if (Directory.GetFiles(modPath).Length == 0)
                        {
                            string[] paths = Directory.GetDirectories(modPath);
                            if (paths.Length == 1)
                            {
                                modPath = paths[0];
                            }
                            else
                            {
                                Console.Clear();
                                Console.WriteLine(localizations[lang]["wrong_mod"]);
                                Console.WriteLine();
                                continue;
                            }
                        }
                        FileSystem.MoveDirectory(modPath, Path.Combine(modDir, new DirectoryInfo(modPath).Name), true);
                        if (Directory.Exists(name))
                            Directory.Delete(name, true);
                        Console.Clear();
                        Console.WriteLine(localizations[lang]["installed_mod"]);
                        Console.WriteLine();
                        array.RemoveAt(index - 1);
                        if (array.Count == 0)
                            return;
                    }
                    else
                    {
                        Console.WriteLine(localizations[lang]["wrong_input"]);
                        Console.WriteLine();
                    }
                }
            }
        }

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public static bool HasNetworkConnection()
        {
            try
            {
                Ping ping = new Ping();
                var result = ping.Send("8.8.8.8", 3000).Status == IPStatus.Success;
                ping.Dispose();
                return result;
            }
            catch (Exception)
            {
            }
            return false;
        }
    }
}
