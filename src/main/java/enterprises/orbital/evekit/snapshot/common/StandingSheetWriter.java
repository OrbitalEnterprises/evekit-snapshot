package enterprises.orbital.evekit.snapshot.common;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.common.Standing;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class StandingSheetWriter {

  // Singleton
  private StandingSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Standings.csv
    // StandingsMeta.csv
    stream.putNextEntry(new ZipEntry("Standings.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Standing Entity", "From ID", "From Name", "Standing");
    List<Long> metaIDs = new ArrayList<Long>();
    long contid = -1;
    List<Standing> batch = Standing.getAllStandings(acct, at, 1000, contid);

    while (batch.size() > 0) {

      for (Standing next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getStandingEntity(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getFromID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getFromName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getStanding(), SheetUtils.CellFormat.DOUBLE_STYLE)); 
        // @formatter:on
        metaIDs.add(next.getCid());
      }

      contid = batch.get(batch.size() - 1).getCid();
      batch = Standing.getAllStandings(acct, at, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("StandingsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "Standing");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
