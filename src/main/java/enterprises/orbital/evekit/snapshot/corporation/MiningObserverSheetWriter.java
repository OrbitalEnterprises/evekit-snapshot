package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.MiningObserver;
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

public class MiningObserverSheetWriter {

  // Singleton
  private MiningObserverSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // MiningObserver.csv
    // MiningObserverMeta.csv
    stream.putNextEntry(new ZipEntry("MiningObserver.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Observer ID", "Observer Type", "Last Updated (Raw)", "Last Updated");
    List<MiningObserver> points = CachedData.retrieveAll(at,
                                                         (contid, at1) -> MiningObserver.accessQuery(acct, contid, 1000,
                                                                                                     false,
                                                                                                     at1,
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any()));

    for (MiningObserver next : points) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getObserverID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getObserverType(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getLastUpdated(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getLastUpdated()), SheetUtils.CellFormat.DATE_STYLE));
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MiningObserverMeta.csv", stream, false, null);
    for (MiningObserver next : points) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "MiningObserver");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

  }

}
