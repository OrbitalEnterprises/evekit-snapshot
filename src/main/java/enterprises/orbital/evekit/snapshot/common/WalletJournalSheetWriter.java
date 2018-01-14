package enterprises.orbital.evekit.snapshot.common;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.common.WalletJournal;
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
    output.printRecord("ID", "Division", "Ref ID", "Date (Raw)", "Date", "Ref Type", "First Party ID",
                       "First Party Type", "Second Party ID", "Second Party Type", "Arg Name 1", "Arg ID 1", "Amount",
                       "Balance", "Reason", "Tax Receiver ID", "Tax Amount", "Location ID", "Transaction ID",
                       "NPC Name", "NPC ID", "Destroyed Ship Type ID", "Character ID", "Corporation ID", "Alliance ID",
                       "Job ID", "Contract ID", "System ID", "PlanetID");

    List<Long> metaIDs = new ArrayList<>();
    long contid = -1;
    List<WalletJournal> batch = WalletJournal.getAllForward(acct, at, 1000, contid);

    while (batch.size() > 0) {

      for (WalletJournal next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getDivision(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getRefID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getDate()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getRefType(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getFirstPartyID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getFirstPartyType(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getSecondPartyID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getSecondPartyType(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getArgName1(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getArgID1(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getAmount(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                   new DumpCell(next.getBalance(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                   new DumpCell(next.getReason(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getTaxReceiverID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getTaxAmount(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                   new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getTransactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getNpcName(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getNpcID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getDestroyedShipTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getAllianceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getJobID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getContractID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getPlanetID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
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
