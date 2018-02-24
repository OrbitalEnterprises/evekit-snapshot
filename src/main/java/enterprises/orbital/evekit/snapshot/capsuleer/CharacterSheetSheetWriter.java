package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterSheet;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class CharacterSheetSheetWriter {

  // Singleton
  private CharacterSheetSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at)
    throws IOException {
    // Sections:
    // CharacterSheet.csv
    // CharacterSheetMeta.csv
    stream.putNextEntry(new ZipEntry("CharacterSheet.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Character ID", "Character Name", "Corporation ID", "RaceID", "DoB (Raw)", "DoB", "BloodlineID",
                       "AncestryID", "Gender", "Alliance ID", "Faction ID", "Description", "Security Status");
    CharacterSheet csheet = CharacterSheet.get(acct, at);
    if (csheet != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(csheet.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getRaceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getDoB(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(csheet.getDoB()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(csheet.getBloodlineID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getAncestryID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getGender(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getAllianceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getFactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getDescription(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(csheet.getSecurityStatus(), SheetUtils.CellFormat.DOUBLE_STYLE));
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("CharacterSheetMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, csheet.getCid(), "CharacterSheet");
    }
    output.flush();
    stream.closeEntry();
  }

}
