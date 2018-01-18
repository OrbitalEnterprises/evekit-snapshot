package enterprises.orbital.evekit.snapshot.corporation;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.corporation.ContainerLog;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

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
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Log Time (Raw)", "Log Time", "Action", "Actor ID", "Actor Name", "Flag", "Item ID", "Item Type ID", "Location ID",
                       "New Configuration", "Old Configuration", "Password Type", "Quantity", "Type ID");
    List<Long> metaIDs = new ArrayList<Long>();
    long contid = -1;
    List<ContainerLog> batch = ContainerLog.getAllForward(acct, at, 1000, contid);

    while (batch.size() > 0) {

      for (ContainerLog next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getLogTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getLogTime()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getAction(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getActorID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getActorName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getFlag(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getItemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getItemTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getNewConfiguration(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getOldConfiguration(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getPasswordType(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
        // @formatter:on
        metaIDs.add(next.getCid());
      }

      contid = batch.get(batch.size() - 1).getLogTime();
      batch = ContainerLog.getAllForward(acct, at, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("ContainerLogsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "ContainerLog");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
