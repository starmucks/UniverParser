using System;
using System.Collections.Generic;

namespace ConsoleApplication4
{
    class Univer
    {
        public string Name { get; set; }

        public string Site { get; set; }

        public string Form { get; set; }

        public string Address { get; set; }

        public string Telephone { get; set; }

        public string Email { get; set; }

        public DateTime LastModified { get; set; }

        public Management Management { get; set; }

        public IDictionary<string, string> Link { get; set; }
    }
}
