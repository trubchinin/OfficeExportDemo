using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.EntityFrameworkCore;
using DataAccessLayer;
using DomainTables;
using ReportPlugin.Abstractions;

namespace OfficeExportDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Підключення до бази даних (SQLite)
            var options = new DbContextOptionsBuilder<LibraryContext>()
                .UseSqlite("Data Source=Solution20/DataAccessLayer/library.db")
                .Options;

            // Отримуємо дані читачів з бази даних
            using var ctx = new LibraryContext(options);
            var repo = new LibraryRepository(options);

            var readers = repo.GetAllReaders()
                .Select(r => (r.Id, $"{r.LastName} {r.FirstName} {r.Patronymic}", r.Category))
                .ToList();

            // Якщо у нас немає даних у базі, додаємо тестові дані
            if (readers.Count == 0)
            {
                Console.WriteLine("База даних порожня. Додаємо тестові дані.");
                readers = new List<(int Id, string FullName, string Category)>
                {
                    (1, "Іваненко Іван Іванович", "студент"),
                    (2, "Петренко Петро Петрович", "аспірант"),
                    (3, "Сидорук Олена Миколаївна", "викладач")
                };
            }
            else
            {
                Console.WriteLine($"Завантажено {readers.Count} читачів з бази даних.");
            }

            // Папка плагінів
            var pluginsDir = Path.Combine(AppContext.BaseDirectory, "Plugins");
            if (!Directory.Exists(pluginsDir))
            {
                Console.WriteLine("Папка Plugins не знайдена.");
                return;
            }

            // Знаходимо всі DLL
            var pluginFiles = Directory.GetFiles(pluginsDir, "*.dll");
            var generators = new List<IReportGenerator>();

            foreach (var file in pluginFiles)
            {
                var asm = Assembly.LoadFrom(file);
                var types = asm.GetTypes()
                    .Where(t => typeof(IReportGenerator).IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract);

                foreach (var type in types)
                    generators.Add((IReportGenerator)Activator.CreateInstance(type)!);
            }

            // Дозволяємо користувачу обрати формат
            Console.WriteLine("Доступні формати звітів:");
            for (int i = 0; i < generators.Count; i++)
                Console.WriteLine($"{i + 1}. {generators[i].FormatName}");

            Console.Write("Оберіть номер формату: ");
            if (!int.TryParse(Console.ReadLine(), out var sel) || sel < 1 || sel > generators.Count)
            {
                Console.WriteLine("Невірний вибір.");
                return;
            }

            var gen = generators[sel - 1];
            var outDir = Path.Combine(AppContext.BaseDirectory, "Reports");
            Directory.CreateDirectory(outDir);

            var ext = gen.FormatName.Equals("Word", StringComparison.OrdinalIgnoreCase)
                ? ".docx" : ".xlsx";

            var outFile = Path.Combine(outDir, $"ReadersReport{ext}");
            gen.Generate(outFile, readers);

            Console.WriteLine($"{gen.FormatName}-звіт збережено: {outFile}");
        }
    }
}
