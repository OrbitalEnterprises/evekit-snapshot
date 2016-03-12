package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.math.BigDecimal;
import java.util.Date;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterSheet;
import enterprises.orbital.evekit.model.character.CharacterSheetBalance;
import enterprises.orbital.evekit.model.character.CharacterSheetClone;
import enterprises.orbital.evekit.model.character.CharacterSheetJump;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class CharacterSheetSheetWriter {

  // Singleton
  private CharacterSheetSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // CharacterSheet.csv
    // CharacterSheetMeta.csv
    stream.putNextEntry(new ZipEntry("CharacterSheet.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Balance", "Character ID", "Character Name", "Corporation ID", "Corporation Name", "Race", "DoB (Raw)", "DoB", "Bloodline",
                       "Ancestry", "Gender", "Alliance Name", "Alliance ID", "Faction Name", "Faction ID", "Clone Name", "Clone Skill Points", "Intelligence",
                       "Memory", "Charisma", "Perception", "Willpower", "Home Station ID", "Clone Jump Date (Raw)", "Clone Jump Date", "Last Respec Date (Raw)",
                       "Last Respec Date", "Last Timed Respec (Raw)", "Last Timed Respec", "Free Respecs", "Free Skill Points", "Clone Type ID",
                       "Remote Station Date (Raw)", "Remote Station Date", "Jump Activation (Raw)", "Jump Activation", "Jump Fatigue (Raw)", "Jump Fatigue",
                       "Jump Last Update (Raw)", "Jump Last Update");
    CharacterSheet csheet = CharacterSheet.get(acct, at);
    if (csheet != null) {
      CharacterSheetBalance bal = CharacterSheetBalance.get(acct, at);
      CharacterSheetClone clone = CharacterSheetClone.get(acct, at);
      CharacterSheetJump jump = CharacterSheetJump.get(acct, at);
      BigDecimal balance = bal.getBalance();
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(csheet.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(balance, balance == null ? SheetUtils.CellFormat.NO_STYLE : SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                 new DumpCell(csheet.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getCorporationName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getRace(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getDoB(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(csheet.getDoB()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(csheet.getBloodline(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getAncestry(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getGender(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getAllianceName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getAllianceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getFactionName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getFactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getIntelligence(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getMemory(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getCharisma(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getPerception(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getWillpower(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getHomeStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(clone.getCloneJumpDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(clone.getCloneJumpDate()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(csheet.getLastRespecDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getLastRespecDate()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(csheet.getLastTimedRespec(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getLastTimedRespec()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(csheet.getFreeRespecs(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getFreeSkillPoints(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(csheet.getRemoteStationDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(csheet.getRemoteStationDate()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(jump.getJumpActivation(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(jump.getJumpActivation()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(jump.getJumpFatigue(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(jump.getJumpFatigue()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(jump.getJumpLastUpdate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(jump.getJumpLastUpdate()), SheetUtils.CellFormat.DATE_STYLE));
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
