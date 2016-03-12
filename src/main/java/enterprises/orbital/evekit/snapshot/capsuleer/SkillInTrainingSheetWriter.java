package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterSkillInTraining;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class SkillInTrainingSheetWriter {

  // Singleton
  private SkillInTrainingSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // SkillInTraining.csv
    // SkillInTrainingMeta.csv
    stream.putNextEntry(new ZipEntry("SkillInTraining.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Skill In Training", "Current Training Queue Time", "Training Start Time (Raw)", "Training Start Time", "Training End Time (Raw)",
                       "Training End Time", "Training Start SP", "Training Destination SP", "Training To Level", "Skill Type ID");
    CharacterSkillInTraining csheet = CharacterSkillInTraining.get(acct, at);
    if (csheet != null) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(csheet.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.isSkillInTraining(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(csheet.getCurrentTrainingQueueTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getTrainingStartTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(csheet.getTrainingStartTime()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(csheet.getTrainingEndTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(csheet.getTrainingEndTime()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(csheet.getTrainingStartSP(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getTrainingDestinationSP(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(csheet.getTrainingToLevel(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
      // @formatter:on
      output.flush();
      stream.closeEntry();

      // Handle MetaData
      output = SheetUtils.prepForMetaData("SkillInTrainingMeta.csv", stream, false, null);
      SheetUtils.dumpNextMetaData(acct, output, csheet.getCid(), "CharacterSkillInTraining");
    }
    output.flush();
    stream.closeEntry();
  }

}
