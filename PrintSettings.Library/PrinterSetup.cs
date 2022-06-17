namespace PrintingProperties.Library;

/// <summary>
/// This class is the printer settings item. contains notes, printername, source, copies etc
/// </summary>
[Serializable]
public class PrinterSetup
{
    public string PrinterNotes = "";
    public PaperSize PaperSize;
    public PaperSource PaperSource;
    public PrinterResolution PrintQuality;
    public string NameOfPrinter = "";
    public bool Landscape = false;
    public bool CanDuplex = false;
    public Duplex DSided;
    public byte[] Devmodearray;
}
