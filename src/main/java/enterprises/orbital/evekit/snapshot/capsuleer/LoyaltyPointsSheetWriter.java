package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.LoyaltyPoints;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class LoyaltyPointsSheetWriter {

  // Singleton
  private LoyaltyPointsSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // LoyaltyPoints.csv
    // LoyaltyPointsMeta.csv
    stream.putNextEntry(new ZipEntry("LoyaltyPoints.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Corporation ID", "Loyalty Points");
    List<LoyaltyPoints> points = CachedData.retrieveAll(at,
                                                        (contid, at1) -> LoyaltyPoints.accessQuery(acct, contid, 1000,
                                                                                                   false,
                                                                                                   at1,
                                                                                                   AttributeSelector.any(),
                                                                                                   AttributeSelector.any()));

    for (LoyaltyPoints next : points) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getLoyaltyPoints(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("LoyaltyPointsMeta.csv", stream, false, null);
    for (LoyaltyPoints next : points) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "LoyaltyPoints");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

  }

}
