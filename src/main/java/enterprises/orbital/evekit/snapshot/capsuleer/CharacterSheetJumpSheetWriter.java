package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterSheetJump;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class CharacterSheetJumpSheetWriter {

  // Singleton
  private CharacterSheetJumpSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at)
    throws IOException {
    // Sections:
    // CharacterSheetJump.csv
    // CharacterSheetJumpMeta.csv
    stream.putNextEntry(new ZipEntry("CharacterSheetJump.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID",
                       "Jump Activation (Raw)", "Jump Activation",
                       "Jump Fatigue (Raw)", "Jump Fatigue",
                       "Jump Last Update (Raw)", "Jump Last Update");
    CharacterSheetJump csheet = CharacterSheetJump.get(acct, at);
    if (csheet != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(csheet.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getJumpActivation(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getJumpActivation()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(csheet.getJumpFatigue(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getJumpFatigue()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(csheet.getJumpLastUpdate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getJumpLastUpdate()), SheetUtils.CellFormat.DATE_STYLE));
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("CharacterSheetJumpMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, csheet.getCid(), "CharacterSheetJump");
    }
    output.flush();
    stream.closeEntry();
  }

}
