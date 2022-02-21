using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace CanadaGames.ViewModels
{
    public class PlacementSummary
    {
        [Display(Name = "Athlete Code")]
        public string ID { get; set; }

        public string Athlete { get; set; }

        public string Contingent { get; set; }

        public string Media_Info { get; set; }

        public double Average { get; set; }

        public int Highest { get; set; }

        public int Lowest { get; set; }

        [Display(Name ="Total Events")]
        public int Total_Events { get; set; }

        [Display(Name = "No. of Sports")]
        public int Number_of_Sports { get; set; }

    }
}
