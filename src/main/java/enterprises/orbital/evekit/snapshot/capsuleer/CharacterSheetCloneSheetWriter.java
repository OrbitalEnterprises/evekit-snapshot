package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterSheetAttributes;
import enterprises.orbital.evekit.model.character.CharacterSheetClone;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class CharacterSheetCloneSheetWriter {

  // Singleton
  private CharacterSheetCloneSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at)
    throws IOException {
    // Sections:
    // CharacterSheetClone.csv
    // CharacterSheetCloneMeta.csv
    stream.putNextEntry(new ZipEntry("CharacterSheetClone.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Clone Jump Date (Raw)", "Clone Jump Date",
                       "Home Station ID", "Home Station Type",
                       "Last Station Change Date (Raw)", "Last Station Change Date");
    CharacterSheetClone csheet = CharacterSheetClone.get(acct, at);
    if (csheet != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(csheet.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getCloneJumpDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getCloneJumpDate()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(csheet.getHomeStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getHomeStationType(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getLastStationChangeDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getLastStationChangeDate()), SheetUtils.CellFormat.DATE_STYLE));
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("CharacterSheetCloneMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, csheet.getCid(), "CharacterSheetClone");
    }
    output.flush();
    stream.closeEntry();
  }

}
