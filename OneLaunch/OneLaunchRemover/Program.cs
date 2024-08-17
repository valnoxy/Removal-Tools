using Microsoft.Win32;
using System.Diagnostics;
using Spectre.Console;

namespace OneLaunchRemover
{
    internal class Program
    {
        private static void Main()
        {
            AnsiConsole.Markup("[red bold]OneLaunch Remover[/] {0}\nCopyright (c) 2018 - 2024 [link=https://valnoxy.dev]valnoxy[/]. All rights reserved.\n\n",
                Markup.Escape("[Version 1.0]"));

            AnsiConsole.Status()
                .Start("Working ...", ctx =>
                {
                    ctx.Spinner(Spinner.Known.Dots);
                    ctx.SpinnerStyle(Style.Parse("green"));

                    // Extract all resources
                    ctx.Status("Killing all processes ...");

                    string[] processNames = ["onelaunch", "onelaunchtray", "chromium"];
                    foreach (var processName in processNames)
                    {
                        KillProcess(processName);
                    }
                    Thread.Sleep(2000);

                    ctx.Status("Removing installation files on all Users ...");
                    foreach (var user in Directory.GetDirectories(@"C:\Users"))
                    {
                        if (user.EndsWith("All Users") || user.EndsWith("Public"))
                        {
                            continue;
                        }

                        AnsiConsole.MarkupLine($"[grey bold]INFO:[/] Current User: '{user}'.");
                        AnsiConsole.MarkupLine("[grey bold]INFO:[/] Removing all files from OneLaunch ...");
                        try
                        {
                            var files = Directory.GetFiles(user, "OneLaunch*.exe", SearchOption.AllDirectories);
                            foreach (var file in files)
                            {
                                RemoveItem(file);
                            }
                        }
                        catch (Exception ex)
                        {
                            AnsiConsole.MarkupLine($"[red bold]ERROR:[/] Error processing directory {user} - {ex.Message}");
                        }

                        AnsiConsole.MarkupLine("[grey bold]INFO:[/] Removing all shortcuts ...");
                        string[] shortcuts = [
                            $@"{user}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\OneLaunch.lnk",
                            $@"{user}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\OneLaunchChromium.lnk",
                            $@"{user}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\OneLaunchUpdater.lnk",
                            $@"{user}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\OneLaunch",
                            $@"{user}\Desktop\OneLaunch.lnk"
                        ];
                        foreach (var shortcut in shortcuts)
                        {
                            RemoveItem(shortcut);
                        }

                        AnsiConsole.MarkupLine("[grey bold]INFO:[/] Removing OneLaunch local user configuration ...");
                        var localPath = $@"{user}\AppData\Local\OneLaunch";
                        if (Directory.Exists(localPath))
                        {
                            try
                            {
                                Directory.Delete(localPath, true);
                            }
                            catch (Exception ex)
                            {
                                AnsiConsole.MarkupLine($"[red bold]ERROR:[/] Failed to remove directory {localPath} - {ex.Message}");
                            }
                        }
                    }

                    ctx.Status("Removing OneLaunch registry entries ...");
                    var sidList = Registry.Users.GetSubKeyNames().Where(sid => !sid.Contains("_Classes"));
                    foreach (var sid in sidList)
                    {
                        RemoveRegistryEntries(sid);
                    }

                    string[] tasks = ["OneLaunchLaunchTask", "ChromiumLaunchTask", "OneLaunchUpdateTask"];
                    foreach (var task in tasks)
                    {
                        RemoveTask($@"C:\Windows\System32\Tasks\{task}");
                    }

                    string[] taskCacheKeys =
                    [
                        @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\TREE\OneLaunchLaunchTask",
                        @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\TREE\ChromiumLaunchTask",
                        @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\TREE\OneLaunchUpdateTask"
                    ];

                    foreach (var taskCacheKey in taskCacheKeys)
                    {
                        RemoveRegistryKey(Registry.LocalMachine, taskCacheKey);
                    }

                    string[] traceCacheKeys =
                    [
                        @"SOFTWARE\Microsoft\Tracing\onelaunch_RASMANCS",
                        @"SOFTWARE\Microsoft\Tracing\onelaunch_RASAPI32"
                    ];

                    foreach (var traceCacheKey in traceCacheKeys)
                    {
                        RemoveRegistryKey(Registry.LocalMachine, traceCacheKey);
                    }
                });
            AnsiConsole.MarkupLine("[bold green]Done[/]: OneLaunch has been successfully removed.");
            AnsiConsole.MarkupLine("[grey]Press any key to exit.[/]");
            Console.ReadLine();
        }

        private static void KillProcess(string processName)
        {
            var processes = Process.GetProcessesByName(processName);
            foreach (var process in processes)
            {
                process.Kill();
                Thread.Sleep(2000);
            }
        }

        private static void RemoveItem(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                    if (File.Exists(path))
                    {
                        AnsiConsole.MarkupLine($"[red bold]ERROR:[/] Failed to remove file: {path}");
                    }
                }
                else if (Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                    if (Directory.Exists(path))
                    {
                        AnsiConsole.MarkupLine($"[red bold]ERROR:[/] Failed to remove directory: {path}");
                    }
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red bold]ERROR:[/] Exception while removing {path}: {ex.Message}");
            }
        }

        private static void RemoveRegistryEntries(string sid)
        {
            var registryPath = $@"{sid}\SOFTWARE\Microsoft\Windows\CurrentVersion\UFH\SHC";
            using (var key = Registry.Users.OpenSubKey(registryPath, true))
            {
                if (key != null)
                {
                    var propertyNames = key.GetValueNames();
                    foreach (var propertyName in propertyNames)
                    {
                        try
                        {
                            key.DeleteValue(propertyName);
                        }
                        catch (Exception ex)
                        {
                            AnsiConsole.MarkupLine($"[red bold]ERROR:[/] Failed to remove OneLaunch: {propertyName} from {registryPath} - {ex.Message}");
                        }
                    }
                }
                else
                {
                    AnsiConsole.MarkupLine($"[yellow bold]WARN:[/] Registry path does not exist: {registryPath}");
                }
            }

            var uninstallKey = $@"{sid}\Software\Microsoft\Windows\CurrentVersion\Uninstall\";
            uninstallKey += "{4947c51a-26a9-4ed0-9a7b-c21e5ae0e71a}_is1";
            RemoveRegistryKey(Registry.Users, uninstallKey);

            string[] runKeys = ["OneLaunch", "OneLaunchChromium", "OneLaunchUpdater"];
            foreach (var key in runKeys)
            {
                RemoveRegistryValue($@"{sid}\Software\Microsoft\Windows\CurrentVersion\Run", key);
            }

            string[] miscKeys = ["OneLaunchHTML_.pdf", "OneLaunch"];
            foreach (var key in miscKeys)
            {
                RemoveRegistryValue($@"{sid}\SOFTWARE\Microsoft\Windows\CurrentVersion\ApplicationAssociationToasts", key);
                RemoveRegistryValue($@"{sid}\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FeatureUsage\AppBadgeUpdated", key);
                RemoveRegistryValue($@"{sid}\SOFTWARE\RegisteredApplications", key);
            }

            string[] paths = [$@"{sid}\Software\OneLaunch", $@"{sid}\SOFTWARE\Classes\OneLaunchHTML"];
            foreach (var path in paths)
            {
                RemoveRegistryKey(Registry.Users, path);
            }
        }

        private static void RemoveRegistryValue(string path, string valueName)
        {
            using var key = Registry.Users.OpenSubKey(path, true);
            if (key?.GetValue(valueName) == null) return;
            
            key.DeleteValue(valueName);
            if (key.GetValue(valueName) == null) return;
            
            AnsiConsole.MarkupLine($"[red bold]ERROR:[/] Failed to remove OneLaunch -> {path}.{valueName}");
        }

        private static void RemoveRegistryKey(RegistryKey baseKey, string subKey)
        {
            try
            {
                baseKey.DeleteSubKeyTree(subKey, false);
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red bold]ERROR:[/] Failed to remove OneLaunch -> {subKey} - {ex.Message}");
            }
        }

        private static void RemoveTask(string taskPath)
        {
            if (!File.Exists(taskPath)) return;
            File.Delete(taskPath);
            
            if (File.Exists(taskPath))
            {
                AnsiConsole.MarkupLine($"[red bold]ERROR:[/] Failed to remove OneLaunch -> {taskPath}");
            }
        }
    }
}
