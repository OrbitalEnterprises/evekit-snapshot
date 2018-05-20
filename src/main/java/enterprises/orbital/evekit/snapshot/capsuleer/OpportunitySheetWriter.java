package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.MiningLedger;
import enterprises.orbital.evekit.model.character.Opportunity;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class OpportunitySheetWriter {

  // Singleton
  private OpportunitySheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // Opportunities.csv
    // OpportunitiesMeta.csv
    stream.putNextEntry(new ZipEntry("Opportunities.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Task ID", "Completed At (Raw)", "Completed At");
    List<Opportunity> points = CachedData.retrieveAll(at,
                                                      (contid, at1) -> Opportunity.accessQuery(acct, contid, 1000,
                                                                                                 false,
                                                                                                 at1,
                                                                                                 AttributeSelector.any(),
                                                                                                 AttributeSelector.any()));

    for (Opportunity next : points) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getTaskID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getCompletedAt(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getCompletedAt()), SheetUtils.CellFormat.DATE_STYLE));
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("OpportunitiesMeta.csv", stream, false, null);
    for (Opportunity next : points) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Opportunity");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

  }

}
