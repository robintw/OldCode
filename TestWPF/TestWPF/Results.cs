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

        public void Add(string _name, string _path)
        {
            this.Add(new Result(_name, _path));
        }
    }
}
