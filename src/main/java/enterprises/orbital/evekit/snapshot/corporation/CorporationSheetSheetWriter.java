package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.corporation.CorporationSheet;
import enterprises.orbital.evekit.model.corporation.MemberLimit;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

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
    // MemberLimit.csv
    // MemberLimitSheetMeta.csv
    stream.putNextEntry(new ZipEntry("CorporationSheet.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Alliance ID", "Ceo ID", "Corporation ID", "Corporation Name", "Description",
                       "Member Count", "Shares", "Station ID", "Tax Rate", "Ticker", "Url",
                       "Date Founded (Raw)", "Date Founded", "Creator ID", "Faction ID", "War Eligible",
                       "px64x64", "px128x128", "px256x256");
    CorporationSheet csheet = CorporationSheet.get(acct, at);
    if (csheet != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(csheet.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getAllianceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getCeoID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getCorporationName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getDescription(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getMemberCount(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getShares(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getTaxRate(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                 new DumpCell(csheet.getTicker(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getUrl(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getDateFounded(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getDateFounded()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(csheet.getCreatorID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getFactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.isWarEligible(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getPx64x64(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getPx128x128(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getPx256x256(), SheetUtils.CellFormat.NO_STYLE));
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("CorporationSheetMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, csheet.getCid(), "CorporationSheet");
    }
    output.flush();
    stream.closeEntry();

    stream.putNextEntry(new ZipEntry("MemberLimit.csv"));
    output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Limit");
    MemberLimit limit = MemberLimit.get(acct, at);
    if (limit != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output,
                                 new DumpCell(limit.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(limit.getMemberLimit(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("MemberLimitMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, limit.getCid(), "MemberLimit");
    }
    output.flush();
    stream.closeEntry();

  }

}
