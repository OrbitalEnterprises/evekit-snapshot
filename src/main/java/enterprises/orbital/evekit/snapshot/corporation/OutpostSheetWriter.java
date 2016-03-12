package enterprises.orbital.evekit.snapshot.corporation;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.corporation.Outpost;
import enterprises.orbital.evekit.model.corporation.OutpostServiceDetail;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class OutpostSheetWriter {

  // Singleton
  private OutpostSheetWriter() {}

  public static List<Long> dumpOutpostServiceDetail(
                                                    SynchronizedEveAccount acct,
                                                    ZipOutputStream stream,
                                                    List<Long> stations,
                                                    long at) throws IOException {
    List<Long> itemIDs = new ArrayList<Long>();
    stream.putNextEntry(new ZipEntry("OutpostServices.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Station ID", "Service Name", "Owner ID", "Min Standing", "Surcharge Per Bad Standing", "Discount Per Good Standing");
    for (long stationID : stations) {
      List<OutpostServiceDetail> batch = OutpostServiceDetail.getAllByStationID(acct, at, stationID);
      if (batch.size() > 0) {
        for (OutpostServiceDetail next : batch) {
          // @formatter:off
          SheetUtils.populateNextRow(output,
                                     new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getServiceName(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getOwnerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getMinStanding(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                     new DumpCell(next.getSurchargePerBadStanding(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                     new DumpCell(next.getDiscountPerGoodStanding(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE));
          // @formatter:on
          itemIDs.add(next.getCid());
        }
        output.println();
      }
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Outposts.csv
    // OutpostsMeta.csv
    // OutpostServices.csv
    // OutpostServicesMeta.csv
    stream.putNextEntry(new ZipEntry("Outposts.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Station ID", "Owner ID", "Station Name", "Solar System ID", "Docking Cost Per Ship Volume", "Office Rental Cost",
                       "Station Type ID", "Reprocessing Efficiency", "Reprocessing Station Take", "Standing Owner ID", "X", "Y", "Z");
    List<Outpost> batch = Outpost.getAll(acct, at);
    List<Long> stationIDs = new ArrayList<Long>();

    // Write out outpost data first
    for (Outpost next : batch) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getOwnerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getStationName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getDockingCostPerShipVolume(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                 new DumpCell(next.getOfficeRentalCost(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                 new DumpCell(next.getStationTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getReprocessingEfficiency(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                 new DumpCell(next.getReprocessingStationTake(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                 new DumpCell(next.getStandingOwnerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getX(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getY(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getZ(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
      // @formatter:on
      stationIDs.add(next.getStationID());
    }
    output.flush();
    stream.closeEntry();

    if (batch.size() > 0) {
      // Wrote at least one outpost so proceed

      // Write out meta data for outposts
      output = SheetUtils.prepForMetaData("OutpostsMeta.csv", stream, false, null);
      for (Outpost next : batch) {
        int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Outpost");
        if (count > 0) output.println();
      }
      output.flush();
      stream.closeEntry();

      // Write out outpost service detail data in the same style as meta data.
      List<Long> writes = dumpOutpostServiceDetail(acct, stream, stationIDs, at);
      if (writes.size() > 0) {
        // Only write out meta-data if an outpost service detail was actually written.
        output = SheetUtils.prepForMetaData("OutpostServicesMeta.csv", stream, false, null);
        for (Long next : writes) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "OutpostServiceDetail");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }
    }
  }

}
