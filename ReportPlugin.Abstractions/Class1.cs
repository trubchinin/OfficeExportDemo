using System.Collections.Generic;

namespace ReportPlugin.Abstractions
{
    public interface IReportGenerator
    {
        /// <summary>
        /// Генерує файл звіту
        /// </summary>
        /// <param name="filePath">Шлях для збереження</param>
        /// <param name="items">Дані</param>
        void Generate(string filePath, List<(int Id, string FullName, string Category)> items);
        
        /// <summary>
        /// Опис формату (наприклад, "Word" або "Excel")
        /// </summary>
        string FormatName { get; }
    }
}
