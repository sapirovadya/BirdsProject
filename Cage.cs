using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BirdsProject1
{
    public class Cage
    {
       public string SerialNumber { get; set; }
        public int Length { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string Material { get; set; }

        public Cage(string SerialNumber, int Length, int Width, int Height, string Material)
        {
            this.SerialNumber = SerialNumber;
            this.Length = Length;
            this.Width = Width;
            this.Height = Height;
            this.Material = Material;
        }

    }
}
