using System;
using System.Reflection;
using Zuby.ADGV;

public class Program {
    public static void Main() {
        var type = typeof(AdvancedDataGridView);
        Console.WriteLine("Properties of AdvancedDataGridView:");
        foreach (var prop in type.GetProperties()) {
            if (prop.Name.Contains("Filter") || prop.Name.Contains("Icon") || prop.Name.Contains("Image")) {
                Console.WriteLine($"- {prop.Name} ({prop.PropertyType.Name})");
            }
        }
    }
}
