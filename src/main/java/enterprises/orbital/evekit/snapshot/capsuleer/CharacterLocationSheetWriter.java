package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterLocation;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class CharacterLocationSheetWriter {

  // Singleton
  private CharacterLocationSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at)
      throws IOException {
    // Sections:
    // CharacterLocation.csv
    // CharacterLocationMeta.csv
    stream.putNextEntry(new ZipEntry("CharacterLocation.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Solar System ID", "Station ID", "Structure ID");
    CharacterLocation csheet = CharacterLocation.get(acct, at);
    if (csheet != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(csheet.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getStructureID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("CharacterLocationMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, csheet.getCid(), "CharacterLocation");
    }
    output.flush();
    stream.closeEntry();
  }

}
