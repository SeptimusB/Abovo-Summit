using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

// Existing credential is read into memory only. Never serialize or log it.
internal sealed class TrialProtection : IDisposable
{
    Action restore;
    internal TrialProtection(Action restore) { this.restore=restore; }
    public void Dispose(){var action=restore;restore=null;if(action!=null)action();}
    internal static string ReadExistingCredential()
    {
        using(var reader=XmlReader.Create(Path.Combine(EngineBenchmark.Repo,"Structure.xml"),
            new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))
        {
            var fields=XDocument.Load(reader).Descendants().Where(e=>e.Name.LocalName=="RejData").ToArray();
            if(fields.Length!=1||String.IsNullOrEmpty(fields[0].Value))throw new InvalidOperationException("Existing protection field unavailable.");
            return fields[0].Value;
        }
    }
}
