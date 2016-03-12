package enterprises.orbital.evekit.snapshot.corporation;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.corporation.CorporationSheet;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class CorporationSheetSheetWriter {

  // Singleton
  private CorporationSheetSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // CorporationSheet.csv
    // CorporationSheetMeta.csv
    stream.putNextEntry(new ZipEntry("CorporationSheet.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Alliance ID", "Alliance Name", "Ceo ID", "Ceo Name", "Corporation ID", "Corporation Name", "Description", "Logo Color 1",
                       "Logo Color 2", "Logo Color 3", "Logo Graphic ID", "Logo Shape 1", "Logo Shape 2", "Logo Shape 3", "Member Count", "Member Limit",
                       "Shares", "Station ID", "Station Name", "Tax Rate", "Ticker", "Url");
    CorporationSheet csheet = CorporationSheet.get(acct, at);
    if (csheet != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(csheet.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getAllianceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getAllianceName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getCeoID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getCeoName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getCorporationName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getDescription(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getLogoColor1(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getLogoColor2(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getLogoColor3(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getLogoGraphicID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getLogoShape1(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getLogoShape2(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getLogoShape3(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getMemberCount(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getMemberLimit(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getShares(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getStationName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getTaxRate(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(csheet.getTicker(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getUrl(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("CorporationSheetMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, csheet.getCid(), "CorporationSheet");
    }
    output.flush();
    stream.closeEntry();
  }

}
