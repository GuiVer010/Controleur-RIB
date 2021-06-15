using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Contrôleur_RIB
{
    class Country// Basic class that serves to hold the required informations about countries, may be used for further data processing if needed
    {
        private String countryCode;
        private String countryName;
        private String countryLocation;
        public Country (String a_countryName, String a_countryCode, String a_countryLocation)
        {
            CountryName = a_countryName;
            CountryCode = a_countryCode;
            CountryLocation = a_countryLocation;
        }
        public String CountryName { get; set; }
        public String CountryCode { get; set; }
        public String CountryLocation { get; set; }
    }
}
