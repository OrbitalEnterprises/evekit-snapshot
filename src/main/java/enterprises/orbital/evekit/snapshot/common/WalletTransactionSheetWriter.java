package enterprises.orbital.evekit.snapshot.common;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.common.WalletTransaction;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

import javax.persistence.Column;

public class WalletTransactionSheetWriter {

  // Singleton
  private WalletTransactionSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at)
    throws IOException {
    // Sections:
    // WalletTransactions.csv
    // WalletTransactionsMeta.csv
    stream.putNextEntry(new ZipEntry("WalletTransactions.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Division", "Transaction ID", "Date (Raw)", "Date", "Quantity", "Type ID",
                       "Price", "Client ID", "Location ID", "Is Buy", "Is Personal", "Journal Transaction ID");

    List<Long> metaIDs = new ArrayList<>();
    long contid = -1;
    List<WalletTransaction> batch = WalletTransaction.getAllForward(acct, at, 1000, contid);

    while (batch.size() > 0) {

      for (WalletTransaction next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getDivision(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getTransactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getDate()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getPrice(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                   new DumpCell(next.getClientID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.isBuy(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.isPersonal(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getJournalTransactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
        // @formatter:on
        if (next.hasMetaData()) metaIDs.add(next.getCid());
      }
      contid = batch.get(batch.size() - 1).getDate();
      batch = WalletTransaction.getAllForward(acct, at, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("WalletTransactionsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "WalletTransaction");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
