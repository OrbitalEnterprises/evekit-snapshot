package enterprises.orbital.evekit.snapshot.common;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.common.AccountStatus;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class AccountStatusSheetWriter {

  // Singleton
  private AccountStatusSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // AccountStatus.csv
    // AccountStatusMeta.csv
    stream.putNextEntry(new ZipEntry("AccountStatus.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Paid Until (Raw)", "Paid Until", "Create Date (Raw)", "Create Date", "Logon Count", "Logon Minutes",
                       "MultiCharacterTraining (Raw)", "MultiCharacterTraining");
    AccountStatus status = AccountStatus.get(acct, at);
    if (status != null) {
      String rawTraining = Arrays.toString(status.getMultiCharacterTraining().toArray());
      StringBuilder trainingDates = new StringBuilder();
      SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss z");
      trainingDates.append('[');
      for (long nextDate : status.getMultiCharacterTraining()) {
        Date when = new Date(nextDate);
        trainingDates.append(formatter.format(when)).append(", ");
      }
      if (status.getMultiCharacterTraining().size() > 0) {
        trainingDates.setLength(trainingDates.length() - 2);
      }
      trainingDates.append(']');
      // @formatter:off
      SheetUtils.populateNextRow(output,
                                 new DumpCell(status.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(status.getPaidUntil(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(status.getPaidUntil()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(status.getCreateDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(status.getCreateDate()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(status.getLogonCount(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(status.getLogonMinutes(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(rawTraining, SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(trainingDates.toString(), SheetUtils.CellFormat.NO_STYLE));
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("AccountStatusMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, status.getCid(), "AccountStatus");
    }
    output.flush();
    stream.closeEntry();
  }
}
