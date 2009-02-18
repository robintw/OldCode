using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;

namespace TestWPF
{
    class Results : ObservableCollection<Result>
    {
        public Results()
        {
            this.Add(new Result("Cooking.pdf", @"D:\Users\Robin Wilson\Documents\Cooking.pdf"));
            this.Add(new Result("BOCP Scan.png", @"D:\Users\Robin Wilson\Documents\BOCP Scan.png"));
        }

        /// <summary>
        /// Adds an item to the list of results
        /// </summary>
        /// <param name="_name">The name of the result</param>
        /// <param name="_path">The path to execute - the icon will also be taken from this path</param>
        public void Add(string _name, string _path)
        {
            this.Add(new Result(_name, _path));
        }
    }
}
