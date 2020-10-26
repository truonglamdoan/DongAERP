using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DongAERP.Areas.Admin.Models
{
    public class CountryModel
    {
        public string CountryTitle { get; set; }
        public string PopulationTitle { get; set; }

        public Country CountrytData { get; set; }
    }
    public class Country
    {
        public string CountryName { get; set; }
        public string Population { get; set; }
    }
}