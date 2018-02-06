package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterShip;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class CharacterShipSheetWriter {

  // Singleton
  private CharacterShipSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at)
      throws IOException {
    // Sections:
    // CharacterShip.csv
    // CharacterShipMeta.csv
    stream.putNextEntry(new ZipEntry("CharacterShip.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Ship Type ID", "Ship Item ID", "Ship Name");
    CharacterShip csheet = CharacterShip.get(acct, at);
    if (csheet != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output,
                                 new DumpCell(csheet.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getShipTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getShipItemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getShipName(), SheetUtils.CellFormat.NO_STYLE));
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("CharacterShipMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, csheet.getCid(), "CharacterShip");
    }
    output.flush();
    stream.closeEntry();
  }

}
