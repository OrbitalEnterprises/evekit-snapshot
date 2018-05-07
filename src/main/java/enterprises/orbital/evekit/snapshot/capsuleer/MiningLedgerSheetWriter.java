package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.MiningLedger;
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

public class MiningLedgerSheetWriter {

  // Singleton
  private MiningLedgerSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // MiningLedger.csv
    // MiningLedgerMeta.csv
    stream.putNextEntry(new ZipEntry("MiningLedger.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Date (Raw)", "Date", "Solar System ID", "Type ID", "Quantity");
    List<MiningLedger> points = CachedData.retrieveAll(at,
                                                       (contid, at1) -> MiningLedger.accessQuery(acct, contid, 1000,
                                                                                                 false,
                                                                                                 at1,
                                                                                                 AttributeSelector.any(),
                                                                                                 AttributeSelector.any(),
                                                                                                 AttributeSelector.any(),
                                                                                                 AttributeSelector.any()));

    for (MiningLedger next : points) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getDate()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MiningLedgerMeta.csv", stream, false, null);
    for (MiningLedger next : points) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "MiningLedger");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

  }

}
