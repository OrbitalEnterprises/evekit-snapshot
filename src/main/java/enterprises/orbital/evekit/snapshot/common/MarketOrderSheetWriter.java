package enterprises.orbital.evekit.snapshot.common;

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
import enterprises.orbital.evekit.model.common.MarketOrder;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class MarketOrderSheetWriter {

  // Singleton
  private MarketOrderSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // MarketOrders.csv
    // MarketOrdersMeta.csv
    stream.putNextEntry(new ZipEntry("MarketOrders.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Order ID", "Wallet Division", "Bid", "Character ID", "Duration", "Escrow", "Issued (Raw)", "Issued", "Min Volume", "Order State",
                       "Price", "Order Range", "Type ID", "Volume Entered", "Volume Remaining", "Region ID", "Location ID", "Is Corp");
    List<Long> metaIDs = new ArrayList<Long>();
    long contid = -1;
    List<MarketOrder> batch = MarketOrder.getAllForward(acct, at, 1000, contid);

    while (batch.size() > 0) {

      for (MarketOrder next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getOrderID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getWalletDivision(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.isBid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getCharID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getDuration(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getEscrow(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                   new DumpCell(next.getIssued(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getIssued()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getMinVolume(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getOrderState(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getPrice(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                   new DumpCell(next.getOrderRange(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getVolEntered(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getVolRemaining(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getRegionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.isCorp(), SheetUtils.CellFormat.NO_STYLE));
        // @formatter:on
        metaIDs.add(next.getCid());
      }

      contid = batch.get(batch.size() - 1).getIssued();
      batch = MarketOrder.getAllForward(acct, at, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MarketOrdersMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "MarketOrder");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
