package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.ContainerLog;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ContainerLogSheetWriter {

  // Singleton
  private ContainerLogSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // ContainerLogs.csv
    // ContainerLogsMeta.csv
    stream.putNextEntry(new ZipEntry("ContainerLogs.csv"));
    final CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Log Time (Raw)", "Log Time", "Action", "Character ID", "Location Flag",
                       "Container ID", "Container Type ID", "Location ID",
                       "New Configuration", "Old Configuration", "Password Type", "Quantity", "Type ID");
    List<Long> metaIDs = new ArrayList<>();

    CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
    CachedData.stream(at, (long contid, AttributeSelector ats) ->
                          ContainerLog.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any()),
                      true, capture)
              .forEach(next -> {
                try {
                  //@formatter:off
                  SheetUtils.populateNextRow(output,
                                             new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getLogTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(new Date(next.getLogTime()), SheetUtils.CellFormat.DATE_STYLE),
                                             new DumpCell(next.getAction(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getLocationFlag(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getContainerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getContainerTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getNewConfiguration(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getOldConfiguration(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getPasswordType(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
                  //@formatter:on
                } catch (IOException e) {
                  capture.handle(e);
                }
                metaIDs.add(next.getCid());
              });

    output.flush();
    stream.closeEntry();

    // Handle MetaData
    CSVPrinter metaOutput = SheetUtils.prepForMetaData("ContainerLogsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, metaOutput, next, "ContainerLog");
      if (count > 0) metaOutput.println();
    }
    metaOutput.flush();
    stream.closeEntry();
  }

}
