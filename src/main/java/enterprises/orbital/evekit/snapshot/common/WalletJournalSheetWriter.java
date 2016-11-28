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
import enterprises.orbital.evekit.model.common.WalletJournal;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class WalletJournalSheetWriter {

  // Singleton
  private WalletJournalSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at)
    throws IOException {
    // Sections:
    // WalletJournal.csv
    // WalletJournalMeta.csv
    stream.putNextEntry(new ZipEntry("WalletJournal.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Account Key", "Ref ID", "Date (Raw)", "Date", "Ref Type ID", "Owner Name 1", "Owner ID 1", "Owner Name 2", "Owner ID 2",
                       "Arg Name 1", "Arg ID 1", "Amount", "Balance", "Reason", "Tax Receiver ID", "Tax Amount", "Owner 1 Type ID", "Owner 2 Type ID");
    List<Long> metaIDs = new ArrayList<Long>();
    long contid = -1;
    List<WalletJournal> batch = WalletJournal.getAllForward(acct, at, 1000, contid);

    while (batch.size() > 0) {

      for (WalletJournal next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getAccountKey(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getRefID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getDate()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getRefTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getOwnerName1(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getOwnerID1(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getOwnerName2(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getOwnerID2(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getArgName1(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getArgID1(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getAmount(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                   new DumpCell(next.getBalance(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                   new DumpCell(next.getReason(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getTaxReceiverID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getTaxAmount(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                   new DumpCell(next.getOwner1TypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getOwner2TypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
        // @formatter:on
        if (next.hasMetaData()) metaIDs.add(next.getCid());
      }
      contid = batch.get(batch.size() - 1).getDate();
      batch = WalletJournal.getAllForward(acct, at, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("WalletJournalMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "WalletJournal");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
