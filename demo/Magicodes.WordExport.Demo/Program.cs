using System;
using System.Windows;

namespace Magicodes.WordExport.Demo;

static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "--verify")
        {
            Verify.Run();
            return;
        }

        var app = new Application();
        app.Run(new MainWindow());
    }
}