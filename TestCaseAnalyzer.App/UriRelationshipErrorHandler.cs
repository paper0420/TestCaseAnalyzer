using DocumentFormat.OpenXml.Packaging;
using System;

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
