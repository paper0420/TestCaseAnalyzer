using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCaseAnalyzer.App
{
    public class UriRelationshipErrorHandler : RelationshipErrorHandler
    {
        public override string Rewrite(Uri partUri, string id, string uri)
        {
            return "http://link-invalido";
        }
    }

}
