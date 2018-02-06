package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterOnline;
import enterprises.orbital.evekit.model.character.CharacterShip;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class CharacterOnlineSheetWriter {

  // Singleton
  private CharacterOnlineSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at)
      throws IOException {
    // Sections:
    // CharacterOnline.csv
    // CharacterOnlineMeta.csv
    stream.putNextEntry(new ZipEntry("CharacterOnline.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Online", "Last Login (Raw)", "Last Login", "Last Logout (Raw)", "Last Logout", "Logins");
    CharacterOnline csheet = CharacterOnline.get(acct, at);
    if (csheet != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output,
                                 new DumpCell(csheet.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.isOnline(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getLastLogin(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getLastLogin()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(csheet.getLastLogout(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getLastLogout()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(csheet.getLogins(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("CharacterOnlineMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, csheet.getCid(), "CharacterOnline");
    }
    output.flush();
    stream.closeEntry();
  }

}
