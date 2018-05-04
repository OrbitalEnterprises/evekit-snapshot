package enterprises.orbital.evekit.snapshot.common;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
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
    final CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Division", "Ref ID", "Date (Raw)", "Date", "Ref Type", "First Party ID",
                       "Second Party ID", "Arg Name 1", "Arg ID 1", "Amount",
                       "Balance", "Reason", "Tax Receiver ID", "Tax Amount", "Context ID", "Context Type",
                       "Description");

    List<Long> metaIDs = new ArrayList<>();
    CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
    CachedData.stream(at, (contid, at1) -> WalletJournal.accessQuery(acct, contid, 1000, false, at1,
                                                                     AttributeSelector.any(), AttributeSelector.any(),
                                                                     AttributeSelector.any(),
                                                                     AttributeSelector.any(), AttributeSelector.any(),
                                                                     AttributeSelector.any(),
                                                                     AttributeSelector.any(), AttributeSelector.any(),
                                                                     AttributeSelector.any(),
                                                                     AttributeSelector.any(), AttributeSelector.any(),
                                                                     AttributeSelector.any(),
                                                                     AttributeSelector.any(), AttributeSelector.any(),
                                                                     AttributeSelector.any(),
                                                                     AttributeSelector.any()), true, capture)
              .forEach(next -> {
                try {
                  // @formatter:off
                  SheetUtils.populateNextRow(output,
                                             new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getDivision(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getRefID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(new Date(next.getDate()), SheetUtils.CellFormat.DATE_STYLE),
                                             new DumpCell(next.getRefType(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getFirstPartyID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getSecondPartyID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getArgName1(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getArgID1(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getAmount(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                             new DumpCell(next.getBalance(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                             new DumpCell(next.getReason(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getTaxReceiverID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getTaxAmount(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                             new DumpCell(next.getContextID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getContextType(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getDescription(), SheetUtils.CellFormat.NO_STYLE));
                  // @formatter:on
                  if (next.hasMetaData()) metaIDs.add(next.getCid());
                } catch (IOException e) {
                  capture.handle(e);
                }
              });

    output.flush();
    stream.closeEntry();

    // Handle MetaData
    CSVPrinter metaOutput = SheetUtils.prepForMetaData("WalletJournalMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, metaOutput, next, "WalletJournal");
      if (count > 0) metaOutput.println();
    }
    metaOutput.flush();
    stream.closeEntry();
  }

}
