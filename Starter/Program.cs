// See https://aka.ms/new-console-template for more information
using System;
using System.IO;
using System.IO.Compression; // Для розпакування архівів
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics; // Для запуску провідника
using System.Management.Automation; // Для роботи з PowerShell
using System.Reflection; // Для роботи з PowerShell
using Microsoft.Win32; // Для перевірки реєстру на наявність Excel
using Microsoft.Extensions.Configuration; 



class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        if (IsExcelInstalled())
        {
            // Створюємо конфігурацію для зчитування з файлу appconfig.json
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appconfig.json", optional: false, reloadOnChange: true)
                .Build();

            // Зчитуємо налаштування з конфігурації
            string nameAppDir = config.GetSection("Settings.AppDirAll").Value;
            if (string.IsNullOrEmpty(nameAppDir)) // Перевіряємо, чи налаштування не порожнє
            {
                MessageBox.Show("Не вдалося знайти налаштування AppDirAll у файлі конфігурації.", 
                    "Помилка конфігурації",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }
            string workFile = "WorkFile";
            string workedDrive = SelectDrive();
            CreateDirectories(workedDrive, nameAppDir);

            UnpackageArchive(
                Path.Combine(workedDrive, nameAppDir, "Reports"),
                Path.Combine(workedDrive, nameAppDir, "Backups"),
                Path.Combine(workedDrive, nameAppDir, workFile, "StarterShootCounter"),
                Path.Combine(workedDrive, nameAppDir, workFile)
            );
        } else
        {
            MessageBox.Show("Microsoft Excel не встановлено на цьому комп'ютері.\n" +
                "Будь ласка, встановіть Microsoft Excel та повторіть спробу.", 
                "Excel не знайдено",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return;
        }
        
    }

    protected static void UnpackageArchive(
        string pathReport, 
        string pathBackups, 
        string thisDirectory,
        string workFile
    )
        // Розпаковує архів з основним файлом та допоміжними файлами
    {
        string myPath = AppDomain.CurrentDomain.BaseDirectory;
        string zipName = "packed.zip";
        // Отримуємо батьківську директорію
        DirectoryInfo parentDirectory = Directory.GetParent(
            Directory.GetParent(myPath).FullName
        );
        // Якщо батьківська директорія існує, продовжуємо
        if (parentDirectory != null)
        {
            string toZip = Path.Combine(parentDirectory.FullName, zipName); // Шлях до архіву
            string pathDestination = workFile;
            // Перевіряємо, чи існує архів
            if (File.Exists(toZip))
            {
                Console.WriteLine("Архів знайдено.");
                // Перевіряємо, чи є в директорії робочий файл
                // Якщо директорія порожня, розпаковуємо архів
                if (GetFirstFile(pathDestination) == "")
                {
                    ZipFile.ExtractToDirectory(toZip, pathDestination);
                    Console.WriteLine(
                        $"Архів розпаковано в директорію: {pathDestination}.", pathDestination
                    );
                }
                
            } else
            {
                if (GetFirstFile(pathDestination) == "")
                {
                    MessageBox.Show(
                        $"Архіву не знайдено за цією адресою: {toZip}.\nВидобуття не відбулося.\nВиберіть файл.", 
                        toZip
                    );
                
                    // Відкриваємо провідник для вибору файлу
                    string openedFile = OpenExplorerAndGetFile(
                        "Архіви zip (*.zip)|*.zip|All files (*.*)|*.*"
                    );
                    // Якщо файл вибрано, розпаковуємо його
                    if (openedFile != "")
                    {
                        ZipFile.ExtractToDirectory(openedFile, pathDestination);
                        Console.WriteLine(
                            $"Архів розпаковано в директорію: {pathDestination}.", pathDestination
                        );
                    }
                    else
                    {
                        Console.WriteLine("Вибір файлу не відбувся. Завершення роботи програми.");
                    }
                } else
                {
                    Console.WriteLine(
                        "Робочий файл вже існує. Видобуття не відбулося.\nПродовження розпаковки."
                    );
                }
            }
            // Викликаємо PowerShell скрипт для зміни макросів
            ExecutePowerShellScript(GetFirstFile(pathDestination), pathReport, pathBackups);

            // Копіюємо себе в директорію з робочим файлом
            SelfCopy(parentDirectory.FullName, thisDirectory, true);
        }
    }

    protected static void CreateDirectories(string drive, string nameAppDir)
    // Створює директорії для звітів, бекапів та самого файлу
    {
        Console.WriteLine("Перевірка наявності директорій.");
        string[] dirs = [
            Path.Combine(drive, nameAppDir, "Reports"),
            Path.Combine(drive, nameAppDir, "Backups"),
            Path.Combine(drive, nameAppDir, "WorkFile"),
            Path.Combine(
            drive, nameAppDir, "WorkFile", "StarterShootCounter"
        )
        ];

        // Запускаємо цикл по кожній директорії та перевіряємо її наявність
        // Створюємо директорії, якщо їх немає
        foreach (string dir in dirs)
        {
            if(!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
                Console.WriteLine("Створено шлях: {0}.", dir);
            } else
            {
                Console.WriteLine("Шлях вже існує: {0}.", dir);
            }
        }
    }

    protected static string SelectDrive()
    // Повертає локальний диск для встановлення файлу
    {
        
        DriveInfo[] allDrive = DriveInfo.GetDrives();
        if(allDrive.Length > 1)
        {
            return allDrive[1].Name;
        } else
        {
            return allDrive[0].Name;
        }
    }

    private static string OpenExplorerAndGetFile(string typesFile)
    // Відкриває провідник для вибору файлу. Повертає шлях до вибраного файлу
    {
        Console.WriteLine("Відкриття провідника для вибору файлу.");
        string selectedFilePath = "";

        // Ініціалізація діалогу вибору файлу
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            Console.WriteLine("Ініціалізація діалогу вибору файлу.");
            // Налаштування діалогу
            openFileDialog.InitialDirectory = Environment.GetFolderPath(
                Environment.SpecialFolder.Desktop
            );
            openFileDialog.Title = "Вкажіть шлях до архіву з файлами обробки витрат ВП";

            // Встановка фільтрів для типів файлів
            openFileDialog.Filter = typesFile;

            // Показ діалогу та отримання результату. Якщо файл вибрано, зберігаємо шлях до нього
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Отримання шляху до вибраного файлу
                selectedFilePath = openFileDialog.FileName;
                Console.WriteLine($"Ви обрали файл: {selectedFilePath}", selectedFilePath);
            }
            else
            {
                MessageBox.Show("Ви не обрали файл. Завершення роботи програми.");
            }

        }
        return selectedFilePath;
    }

    private static void CallPowerShellScript(string scriptPath, string fileExcelPath)
    // Викликає PowerShell скрипт. Для роботи з COM об'єктами
    // Не використовується
    {
        // Перевірка, чи існує файл
        if (!File.Exists(fileExcelPath))
        {
            Console.WriteLine("Робочого файлу для обробки не знайдено. (CallPowerShellScript)");
            return;
        }
        if(!File.Exists(scriptPath))
        {
            Console.WriteLine("Скрипт не знайдено.");
            return;
        } else
        {
            // Створення нового процесу
            Process process = new Process();

            // Запуск скрипта
            process.StartInfo.FileName = "powershell.exe";

            // -ExecutionPolicy ByPass обходить політику виконання скриптів
            // -File вказує, що буде запускатися файл
            process.StartInfo.Arguments = $"-ExecutionPolicy ByPass -File \"{scriptPath}\" -PathFileExcel \"{fileExcelPath}\"";

            // Не показувати вікно консолі
            process.StartInfo.CreateNoWindow = true;

            // Не використовувати оболонку ОС
            process.StartInfo.UseShellExecute = false;

            try
            {
                Console.WriteLine("Запуск скрипту зміни макросів.");

                // Запуск процесу
                process.Start();

                // Очікування завершення процесу
                process.WaitForExit();

                Console.WriteLine("Скрипт завершив роботу.");
                Console.WriteLine($"Код виходу: {process.ExitCode}");
            } catch (Exception ex)
            {
                Console.WriteLine($"Сталася помилка: {ex.Message}");
            }

        }
    }

    private static string GetFirstFile(string pathToFile)
    // Повертає перший файл з директорії. Якщо файлів немає, повертає порожній рядок
    {
        // Повертає файл з директорії
        DirectoryInfo dirInfo = new DirectoryInfo(pathToFile);
        FileInfo[] files = dirInfo.GetFiles();
        if (files.Length > 0)
        {
            return files[0].FullName;
        }
        else
        {
            Console.WriteLine("Робочого файлу у директорії не знайдено. (GetFirstFile)");
            return "";
        }
    }

    public static void ExecutePowerShellScript(
        string workFilePath, string pathReport, string pathBackups
        )
    // Виконує PowerShell скрипт для зміни макросів
    {
        // Завантажуємо скрипт з вбудованого ресурсу
        var assembly = Assembly.GetExecutingAssembly();
        var resourseName = "Starter.edit.ps1";
        string scriptContent = "";

        // Читаємо вміст вбудованого ресурсу за допомогою потоку
        using (Stream stream = assembly.GetManifestResourceStream(resourseName))
        {
            // Перевіряємо, чи вдалося знайти ресурс. Якщо потік null, ресурс не знайдено
            if (stream == null)
            {
                Console.WriteLine("Не вдалося знайти вбудований ресурс.");
                return;
            }
            using (StreamReader reader = new StreamReader(stream))
            {
                scriptContent = reader.ReadToEnd();
            }
        }

        // Виконуємо скрипт
        using (PowerShell ps = PowerShell.Create())
        {
            ps.AddScript(scriptContent); // Додаємо скрипт
            ps.AddParameter("PathFileExcel", workFilePath); // Додаємо параметри 
            ps.AddParameter("PathToReport", pathReport); // Додаємо параметри
            ps.AddParameter("PathToBackup", pathBackups); // Додаємо параметри
            ps.AddParameter("ExecutionPolicy", "ByPass"); // Обхід політики виконання
            var result = ps.Invoke(); // Виконуємо скрипт

            // Якщо є помилки, виводимо їх
            if (ps.Streams.Error.Any())
            {
                Console.WriteLine($"Сталася помилка під час виконання скрипта.");
                foreach(var error in ps.Streams.Error) 
                // Виводимо помилки
                {
                    Console.WriteLine(error.ToString());
                }
            }
        }
        Console.WriteLine("Макроси успішно змінено.");
    }

    public static void SelfCopy(string parentDirectory, string destinationDir, bool recursive)
    // Копіює себе в директорію з робочим файлом
    {
        var dir = new DirectoryInfo(parentDirectory); // Отримуємо інформацію про батьківську директорію

        // Перевіряємо, чи існує батьківська директорія
        if (!dir.Exists)
        {
            // Якщо батьківська директорія не існує, викидаємо виключення
            throw new DirectoryNotFoundException(
                $"Батьківська директорія не існує або не знайдена: {dir.FullName}"
            );
        }

        // Якщо директорія призначення не існує, створюємо її
        if (!Directory.Exists(destinationDir))
        {
            Directory.CreateDirectory(destinationDir);
        }

        foreach (FileInfo file in dir.GetFiles())
        // Копіюємо кожен файл з батьківської директорії в директорію призначення
        {
            string targetFilePath = Path.Combine(destinationDir, file.Name);
            file.CopyTo(targetFilePath, true);
        }

        if (recursive) // Якщо потрібно копіювати піддиректорії
        {
            // Рекурсивно копіюємо всі піддиректорії
            foreach (DirectoryInfo subDir in dir.GetDirectories())
            {
                string newDestinationDir = Path.Combine(destinationDir, subDir.Name);
                SelfCopy(subDir.FullName, newDestinationDir, true); // Рекурсивний виклик
            }
        }
        Console.WriteLine($"Файли було скопійовано в {destinationDir}");
    }

    protected static bool IsExcelInstalled()
    // Перевіряє, чи встановлений Excel на комп'ютері
    {
        RegistryKey keyExcel = Registry.ClassesRoot.OpenSubKey(@"Excel.Application");
        if (keyExcel != null)
        {
            keyExcel.Close();
            return true; // Excel встановлений
        } else
        {
            return false; // Excel не встановлений
        }
    }
}