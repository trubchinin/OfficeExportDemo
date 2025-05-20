using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ReportPlugin.Abstractions;

namespace OfficeExportDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Використовуємо тестові дані замість бази даних
                var readers = new List<(int Id, string FullName, string Category)>
                {
                    (1, "Іваненко Іван Іванович", "студент"),
                    (2, "Петренко Петро Петрович", "аспірант"),
                    (3, "Сидорук Олена Миколаївна", "викладач")
                };
                
                Console.WriteLine($"Підготовлено {readers.Count} тестових записів для звіту.");

                // Папка плагінів - шукаємо відносно поточної директорії
                var currentDir = Directory.GetCurrentDirectory();
                Console.WriteLine($"Поточна директорія: {currentDir}");
                
                var pluginsDir = Path.Combine(currentDir, "Plugins");
                Console.WriteLine($"Шукаємо плагіни в: {pluginsDir}");
                
                if (!Directory.Exists(pluginsDir))
                {
                    Console.WriteLine("Папка Plugins не знайдена. Спробую використати bin/Debug/net9.0/Plugins");
                    pluginsDir = Path.Combine(currentDir, "bin", "Debug", "net9.0", "Plugins");
                    Console.WriteLine($"Шукаємо плагіни в: {pluginsDir}");
                    
                    if (!Directory.Exists(pluginsDir))
                    {
                        Console.WriteLine("Папка Plugins не знайдена. Завершення роботи.");
                        return;
                    }
                }

                // Знаходимо всі DLL
                var pluginFiles = Directory.GetFiles(pluginsDir, "*.dll");
                Console.WriteLine($"Знайдено {pluginFiles.Length} DLL файлів");
                
                var generators = new List<object>();

                // Шукаємо типи, що реалізують інтерфейс IReportGenerator
                foreach (var file in pluginFiles)
                {
                    try
                    {
                        Console.WriteLine($"Спроба завантажити: {Path.GetFileName(file)}");
                        var asm = Assembly.LoadFrom(file);
                        var types = asm.GetTypes()
                            .Where(t => !t.IsInterface && !t.IsAbstract && t.GetInterfaces().Any(i => i.FullName?.EndsWith(".IReportGenerator") == true));

                        foreach (var type in types)
                        {
                            try
                            {
                                var instance = Activator.CreateInstance(type);
                                if (instance != null)
                                {
                                    generators.Add(instance);
                                    Console.WriteLine($"Завантажено плагін: {type.FullName}");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Помилка при створенні екземпляру {type.FullName}: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Помилка при завантаженні {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                if (generators.Count == 0)
                {
                    Console.WriteLine("Не знайдено жодного плагіна.");
                    return;
                }

                // Дозволяємо користувачу обрати формат
                Console.WriteLine("\nДоступні формати звітів:");
                for (int i = 0; i < generators.Count; i++)
                {
                    var pluginFormat = generators[i].GetType().GetProperty("FormatName")?.GetValue(generators[i], null) as string;
                    Console.WriteLine($"{i + 1}. {pluginFormat}");
                }

                Console.Write("\nОберіть номер формату (або Enter для виходу): ");
                var input = Console.ReadLine();
                if (string.IsNullOrEmpty(input) || !int.TryParse(input, out var sel) || sel < 1 || sel > generators.Count)
                {
                    Console.WriteLine("Невірний вибір. Завершення роботи.");
                    return;
                }

                var gen = generators[sel - 1];
                var outDir = Path.Combine(currentDir, "Reports");
                Directory.CreateDirectory(outDir);

                // Отримуємо значення форматного імені через рефлексію
                var formatNameProperty = gen.GetType().GetProperty("FormatName");
                var formatName = formatNameProperty?.GetValue(gen, null) as string;

                var ext = formatName?.Equals("Word", StringComparison.OrdinalIgnoreCase) == true
                    ? ".docx" : ".xlsx";

                var outFile = Path.Combine(outDir, $"ReadersReport{ext}");
                
                // Викликаємо Generate через рефлексію
                var generateMethod = gen.GetType().GetMethod("Generate");
                generateMethod?.Invoke(gen, new object[] { outFile, readers });

                Console.WriteLine($"{formatName}-звіт збережено: {outFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Помилка: {ex.Message}");
                if (ex.InnerException != null)
                    Console.WriteLine($"Внутрішня помилка: {ex.InnerException.Message}");
            }
            
            Console.WriteLine("\nНатисніть Enter для завершення...");
            Console.ReadLine();
        }
    }
}
