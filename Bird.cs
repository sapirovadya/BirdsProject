using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BirdsProject1
{
    public class Bird
    {
        public int SerialNumber { get; set; }
        public string species { get; set; }
        public string subSpecies { get; set; }
        public DateTime hatchDate { get; set; }
        public string gender { get; set; }
        public string cageNumber { get; set; }
        public string SerialNumberMother { get; set; }
        public string SerialNumberfather { get; set; }

        public Bird(int SerialNumber, string species,string subSpecies, DateTime hatchDate, string gender, string cageNumber, string SerialNumberMother, string SerialNumberfather)
        {
            this.SerialNumber = SerialNumber;
            this.species = species;
            this.subSpecies = subSpecies;
            this.hatchDate = hatchDate;
            this.gender = gender;
            this.cageNumber = cageNumber;
            this.SerialNumberMother = SerialNumberMother;
            this.SerialNumberfather = SerialNumberfather;

        }



    }



}
