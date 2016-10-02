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
import enterprises.orbital.evekit.model.common.Location;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class LocationSheetWriter {

  // Singleton
  private LocationSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at)
    throws IOException {
    // Sections:
    // Locations.csv
    // LocationsMeta.csv
    stream.putNextEntry(new ZipEntry("Locations.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Item ID", "Item Name", "X", "Y", "Z");
    List<Location> locations = new ArrayList<Location>();
    long contid = -1;
    List<Location> batch = Location.getAllLocations(acct, at, 1000, contid);
    while (batch.size() > 0) {
      locations.addAll(batch);
      contid = batch.get(batch.size() - 1).getItemID();
      batch = Location.getAllLocations(acct, at, 1000, contid);
    }

    for (Location next : locations) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getItemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getItemName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getX(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getY(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getZ(), SheetUtils.CellFormat.DOUBLE_STYLE)); 
      // @formatter:on
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("LocationsMeta.csv", stream, false, null);
    for (Location next : locations) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Location");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
