using System;
using System.Reflection;
using System.Drawing;
using Zuby.ADGV;

public class Program {
    public static void Main() {
        var type = typeof(AdvancedDataGridView);
        Console.WriteLine("Static Fields in AdvancedDataGridView:");
        foreach (var field in type.GetFields(BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public)) {
            if (field.FieldType == typeof(Image) || field.FieldType == typeof(Bitmap)) {
                Console.WriteLine($"- {field.Name} ({field.FieldType.Name})");
            }
        }
        
        // Also check ColumnHeaderCell
        var cellType = typeof(DataGridViewColumnHeaderCell);
        // ... or more likely the ADGV specific one
        // Let's search for types in the assembly
        foreach (var t in type.Assembly.GetTypes()) {
            if (t.Name.Contains("ColumnHeaderCell")) {
                Console.WriteLine($"\nProperties/Fields in {t.Name}:");
                foreach (var f in t.GetFields(BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public)) {
                     if (f.FieldType == typeof(Image) || f.FieldType == typeof(Bitmap)) {
                        Console.WriteLine($"- {f.Name} ({f.FieldType.Name})");
                    }
                }
            }
        }
    }
}
