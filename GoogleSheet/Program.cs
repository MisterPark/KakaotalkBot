using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheet
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Settings settings;
            settings = Settings.Load();
            Settings.Save(settings);


        }
    }
}
