using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Business.Handler
{

    //<Rule>
    //  <Type>字符串</Type>
    //  <IsNull>是</IsNull>
    //  <Note></Note>
    //  <Locations>
    //    <Loc>sheetname,1,A，60</Loc>
    //    <Loc>sheetname,1,B，40</Loc>
    //  </Locations>
    //</Rule>
    class RuleModel
    {
        internal List<Rule> Rules { get; set; }
    }

    class Rule
    {
        internal string Type { get; set; }
        internal string IsNull { get; set; }
        internal string Note { get; set; }
        internal List<Location> Locations { get; set; }
    }

    class Location
    {
        internal string loc { get; set; }
    }
}
