// See https://aka.ms/new-console-template for more information
using System;
using System.IO;
using System.IO.Compression; // Для розпакування архівів
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics; // Для запуску провідника


class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        string nameAppDir = "Artilery_3027";
        string workedDrive = SelectDrive();
        CreateDirectories(workedDrive, nameAppDir);
        UnpackageArchive(workedDrive, nameAppDir);
    }

    protected static void UnpackageArchive(string workedDrive, string nameAppDir)
        // Розпаковує архів з основним файлом та допоміжними файлами
    {
        string myPath = AppDomain.CurrentDomain.BaseDirectory;
        string zipName = "packed.zip";
        DirectoryInfo parentDirectory = Directory.GetParent(
            Directory.GetParent(myPath).FullName
        );
        if (parentDirectory != null)
        {
            string toZip = Path.Combine(parentDirectory.FullName, zipName);
            string pathDestination = Path.Combine(workedDrive, nameAppDir, "WorkFile");
            if (File.Exists(toZip))
            {
                Console.WriteLine("Архів знайдено.");
                
                ZipFile.ExtractToDirectory(toZip, pathDestination);
                Console.WriteLine(
                    $"Архів розпаковано в директорію: {pathDestination}.", pathDestination
                );
            } else
            {
                MessageBox.Show(
                    $"Архіву не знайдено за цією адресою: {toZip}.\nВидобуття не відбулося.\nВиберіть файл.", 
                    toZip
                );
                
                // Відкриваємо провідник для вибору файлу
                string openedFile = OpenExplorerAndGetFile(
                    "Архіви zip (*.zip)|*.zip|All files (*.*)|*.*"
                );
                if (openedFile != "")
                {
                    ZipFile.ExtractToDirectory(openedFile, pathDestination);
                    Console.WriteLine(
                        $"Архів розпаковано в директорію: {pathDestination}.", pathDestination
                    );
                } else
                {
                    Console.WriteLine("Вибір файлу не відбувся. Завершення роботи програми.");
                }
            }
            // Виклик PowerShell скрипта для зміни макросів
            CallPowerShellScript(
                Path.Combine(myPath, "edit.ps1"), 
                Path.Combine(pathDestination, GetFirstFile(pathDestination))
            );
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
        ];
        foreach(string dir in dirs)
        {
            if(!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
                Console.WriteLine("Створено шлях: {0}.", dir);
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
    {
        Console.WriteLine("Відкриття провідника для вибору файлу.");
        string selectedFilePath = "";
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
}