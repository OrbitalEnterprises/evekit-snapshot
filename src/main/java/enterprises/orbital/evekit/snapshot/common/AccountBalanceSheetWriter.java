package enterprises.orbital.evekit.snapshot.common;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.common.AccountBalance;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class AccountBalanceSheetWriter {

  // Singleton
  private AccountBalanceSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // AccountBalance.csv
    // AccountBalanceMeta.csv
    stream.putNextEntry(new ZipEntry("AccountBalance.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Division", "Balance");
    List<AccountBalance> balances = SheetUtils.retrieveAll(at, (ctid, tm) -> AccountBalance.accessQuery(acct, ctid, 1000, false, tm, SheetUtils.ANY_SELECTOR, SheetUtils.ANY_SELECTOR));

    for (AccountBalance next : balances) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getDivision(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getBalance(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE));
      // @formatter:on
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("AccountBalanceMeta.csv", stream, false, null);
    for (AccountBalance next : balances) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "AccountBalance");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
