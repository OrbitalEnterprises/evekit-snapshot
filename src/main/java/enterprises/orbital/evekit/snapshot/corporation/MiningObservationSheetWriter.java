package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.MiningObservation;
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

public class MiningObservationSheetWriter {

  // Singleton
  private MiningObservationSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // MiningObservation.csv
    // MiningObservationMeta.csv
    stream.putNextEntry(new ZipEntry("MiningObservation.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Observer ID", "Character ID", "Type ID", "Recorded Corporation ID", "Quantity",
                       "Last Updated (Raw)", "Last Updated");
    List<MiningObservation> points = CachedData.retrieveAll(at,
                                                            (contid, at1) -> MiningObservation.accessQuery(acct, contid,
                                                                                                           1000,
                                                                                                           false,
                                                                                                           at1,
                                                                                                           AttributeSelector.any(),
                                                                                                           AttributeSelector.any(),
                                                                                                           AttributeSelector.any(),
                                                                                                           AttributeSelector.any(),
                                                                                                           AttributeSelector.any(),
                                                                                                           AttributeSelector.any()));

    for (MiningObservation next : points) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getObserverID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getRecordedCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getLastUpdated(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getLastUpdated()), SheetUtils.CellFormat.DATE_STYLE));
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MiningObservationMeta.csv", stream, false, null);
    for (MiningObservation next : points) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "MiningObservation");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

  }

}
