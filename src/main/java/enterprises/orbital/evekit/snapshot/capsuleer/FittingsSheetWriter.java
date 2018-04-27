package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.Fitting;
import enterprises.orbital.evekit.model.character.FittingItem;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class FittingsSheetWriter {

  // Singleton
  private FittingsSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // Fittings.csv
    // FittingsMeta.csv
    // FittingItems.csv
    // FittingItemsMeta.csv
    stream.putNextEntry(new ZipEntry("Fittings.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Fitting ID", "Name", "Description", "Ship Type ID");
    List<Fitting> fittings = CachedData.retrieveAll(at, (contid, at1) -> Fitting.accessQuery(acct, contid, 1000,
                                                                                             false,
                                                                                             at1,
                                                                                             AttributeSelector.any(),
                                                                                             AttributeSelector.any(),
                                                                                             AttributeSelector.any(),
                                                                                             AttributeSelector.any()));

    for (Fitting next : fittings) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getFittingID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getName(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getDescription(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getShipTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("FittingsMeta.csv", stream, false, null);
    for (Fitting next : fittings) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Fitting");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

    stream.putNextEntry(new ZipEntry("FittingItems.csv"));
    output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Fitting ID", "Type ID", "Flag", "Quantity");
    List<FittingItem> items = CachedData.retrieveAll(at, (contid, at12) -> FittingItem.accessQuery(acct,
                                                                                                   contid,
                                                                                                   1000,
                                                                                                   false,
                                                                                                   at12,
                                                                                                   AttributeSelector.any(),
                                                                                                   AttributeSelector.any(),
                                                                                                   AttributeSelector.any(),
                                                                                                   AttributeSelector.any()));

    for (FittingItem next : items) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getFittingID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getFlag(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("FittingItemsMeta.csv", stream, false, null);
    for (FittingItem next : items) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "FittingItem");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

  }

}
