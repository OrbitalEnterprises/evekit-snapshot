package enterprises.orbital.evekit.snapshot.common;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.common.FacWarStats;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class FacWarStatsSheetWriter {

  // Singleton
  private FacWarStatsSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // FacWarStats.csv
    // FacWarStatsMeta.csv
    stream.putNextEntry(new ZipEntry("FacWarStats.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Current Rank", "Enlisted (Raw)", "Enlisted", "Faction ID", "Highest Rank", "Kills Last Week", "Kills Total",
                       "Kills Yesterday", "Pilots", "Victory Points Last Week", "Victory Points Total", "Victory Points Yesterday");
    FacWarStats stats = FacWarStats.get(acct, at);
    if (stats != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(stats.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(stats.getCurrentRank(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(stats.getEnlisted(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(stats.getEnlisted()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(stats.getFactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(stats.getHighestRank(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(stats.getKillsLastWeek(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(stats.getKillsTotal(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(stats.getKillsYesterday(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(stats.getPilots(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(stats.getVictoryPointsLastWeek(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(stats.getVictoryPointsTotal(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(stats.getVictoryPointsYesterday(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("FacWarStatsMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, stats.getCid(), "FacWarStats");
    }
    output.flush();
    stream.closeEntry();
  }
}
